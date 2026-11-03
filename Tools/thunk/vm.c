/* ASF Windows x64 resumable VM. No imports, globals, relocations or CRT.
 * VBA owns every VARIANT and all invocation storage; C borrows it only while
 * running. Host operations yield before VBA/COM can reenter the engine. */
#include <windows.h>
#include <oleauto.h>
#include <stddef.h>
#include <stdint.h>

typedef struct { int32_t op, a, b, c; } Instruction;
typedef struct {
    Instruction *code;
    VARIANT *constants, *stack;
    HRESULT (WINAPI *copy)(VARIANT *, VARIANT *);
    HRESULT (WINAPI *clear)(VARIANT *);
    int32_t count, capacity, pc, sp, request, argument, signal, error;
    int32_t constant_count, reserved;
} Frame;
_Static_assert(sizeof(VARIANT) == 24, "x64 VARIANT ABI");
_Static_assert(sizeof(Instruction) == 16, "instruction ABI");
_Static_assert(offsetof(Frame, count) == 40 && sizeof(Frame) == 80, "frame ABI");

enum { HALT, CONSTANT, EXPR, STMT, STORE, RETURN_VALUE, DROP, JUMP,
       JFALSE, TRUTH, BINARY, UNARY, NONNULL, PUSH_BOOL, SIGNAL, LET, JTRUE };

/* CallWindowProc-compatible entry: frame is the first argument. */
intptr_t vm(Frame *f, uintptr_t a, uintptr_t b, uintptr_t c) {
    (void)a; (void)b; (void)c;
    if (!f || !f->code || !f->stack || !f->constants || !f->copy || !f->clear)
        return -1;
    if (f->signal) {
        if (f->pc < 1 || f->pc > f->count) return -1;
        Instruction *previous = f->code + f->pc - 1;
        int target = f->signal == 1 ? previous->b : f->signal == 2 ? previous->c : -1;
        if (previous->op == STMT && target >= 0) {
            f->pc = target;
            f->signal = 0;
        } else return 0;
    }
    f->request = 0;
    for (;;) {
        if (f->pc < 0 || f->pc >= f->count || f->sp < 0 || f->sp > f->capacity) return -1;
        Instruction *i = f->code + f->pc;
        VARIANT *top = f->sp ? f->stack + f->sp - 1 : 0;
        switch (i->op) {
        case HALT: return 0;
        case CONSTANT:
            if (f->sp >= f->capacity || i->a < 0 || i->a >= f->constant_count) return -1;
            f->error = f->copy(f->stack + f->sp, f->constants + i->a);
            if (f->error < 0) return -2;
            ++f->sp;
            break;
        case DROP:
            if (!top) return -1;
            f->error = f->clear(top);
            if (f->error < 0) return -2;
            --f->sp;
            break;
        case JUMP: f->pc = i->a; continue;
        case JFALSE: case JTRUE: {
            if (!top || V_VT(top) != VT_BOOL) return -1;
            int yes = V_BOOL(top) != 0;
            V_VT(top) = VT_EMPTY;
            --f->sp;
            if (yes == (i->op == JTRUE)) { f->pc = i->a; continue; }
            break;
        }
        case NONNULL:
            if (!top) return -1;
            if (V_VT(top) != VT_EMPTY && V_VT(top) != VT_NULL) { f->pc = i->a; continue; }
            V_VT(top) = VT_EMPTY;
            --f->sp;
            break;
        case PUSH_BOOL:
            if (f->sp >= f->capacity) return -1;
            top = f->stack + f->sp++;
            if (V_VT(top) != VT_EMPTY) return -1;
            V_VT(top) = VT_BOOL; V_BOOL(top) = i->a ? VARIANT_TRUE : VARIANT_FALSE;
            break;
        case SIGNAL: f->signal = i->a; return 0;
        case LET:
            if (!top) return -1;
            if (V_VT(top) == VT_DISPATCH || V_VT(top) == VT_UNKNOWN) goto host;
            break;
        case TRUTH: {
            if (!top) return -1;
            int yes;
            switch (V_VT(top)) {
            case VT_EMPTY: case VT_NULL: yes = 0; break;
            case VT_BOOL: yes = V_BOOL(top) != 0; break;
            case VT_R8: yes = V_R8(top) != 0; break;
            case VT_I2: yes = V_I2(top) != 0; break;
            case VT_I4: yes = V_I4(top) != 0; break;
            default: goto host;
            }
            V_VT(top) = VT_BOOL; V_BOOL(top) = yes ? VARIANT_TRUE : VARIANT_FALSE;
            break;
        }
        case BINARY: {
            if (f->sp < 2) return -1;
            VARIANT *left = top - 1;
            /* Preserve VBA subtype/coercion/overflow behavior via the host for
             * everything except two finite Doubles and these exact operators. */
            if (V_VT(left) != VT_R8 || V_VT(top) != VT_R8) goto host;
            union { double d; uint64_t u; } x, y, z;
            x.d = V_R8(left); y.d = V_R8(top);
            if ((x.u & UINT64_C(0x7ff0000000000000)) == UINT64_C(0x7ff0000000000000) ||
                (y.u & UINT64_C(0x7ff0000000000000)) == UINT64_C(0x7ff0000000000000)) goto host;
            int yes = 0;
            switch (i->a) {
            case 0: z.d = x.d + y.d; break;
            case 1: z.d = x.d - y.d; break;
            case 2: z.d = x.d * y.d; break;
            case 3: if (y.d == 0) goto host; z.d = x.d / y.d; break;
            case 4: yes = x.d < y.d; break;
            case 5: yes = x.d <= y.d; break;
            case 6: yes = x.d > y.d; break;
            case 7: yes = x.d >= y.d; break;
            case 8: yes = x.d == y.d; break;
            case 9: yes = x.d != y.d; break;
            default: goto host;
            }
            if (i->a < 4) {
                if ((z.u & UINT64_C(0x7ff0000000000000)) == UINT64_C(0x7ff0000000000000)) goto host;
                V_R8(left) = z.d;
            } else { V_VT(left) = VT_BOOL; V_BOOL(left) = yes ? VARIANT_TRUE : VARIANT_FALSE; }
            V_VT(top) = VT_EMPTY;
            --f->sp;
            break;
        }
        case EXPR: case STMT: case STORE: case RETURN_VALUE: case UNARY:
            goto host;
        default: return -1;
        }
        ++f->pc;
        continue;
host:
        f->request = i->op; f->argument = i->a;
        return 1;
    }
}
