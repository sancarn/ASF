/* Native control/ABI tests, complementing semantic checks inside Excel. */
#include "vm.c"
#include <assert.h>
#include <stdio.h>

static Frame frame(Instruction *code, int n, VARIANT *pool, int constants, VARIANT *stack, int cap) {
    Frame f = {0};
    f.code = code; f.count = n; f.constants = pool; f.constant_count = constants;
    f.stack = stack; f.capacity = cap; f.copy = VariantCopy; f.clear = VariantClear;
    return f;
}

int main(void) {
    VARIANT pool[3] = {0}, stack[8] = {0};
    V_VT(&pool[0]) = VT_R8; V_R8(&pool[0]) = 7;
    V_VT(&pool[1]) = VT_R8; V_R8(&pool[1]) = 3;
    V_VT(&pool[2]) = VT_BSTR; V_BSTR(&pool[2]) = SysAllocString(L"owned string");
    for (int op = 0; op < 10; ++op) {
        Instruction code[] = {{CONSTANT,0,0,0},{CONSTANT,1,0,0},{BINARY,op,0,0},{HALT,0,0,0}};
        Frame f = frame(code,4,pool,3,stack,8);
        assert(vm(&f,0,0,0) == 0 && f.sp == 1);
        if (op < 4) {
            double expected[] = {10,4,21,7.0/3.0};
            assert(V_VT(stack) == VT_R8 && V_R8(stack) == expected[op]);
        } else {
            int expected[] = {0,0,1,1,0,1};
            assert(V_VT(stack) == VT_BOOL && !!V_BOOL(stack) == expected[op-4]);
        }
        VariantClear(stack);
    }
    {
        Instruction code[] = {{CONSTANT,2,0,0},{TRUTH,0,0,0},{JFALSE,5,0,0},{DROP,0,0,0},{HALT,0,0,0},{HALT,0,0,0}};
        Frame f = frame(code,6,pool,3,stack,8);
        assert(vm(&f,0,0,0) == 1 && f.pc == 1 && f.request == TRUTH);
        assert(V_BSTR(stack) != V_BSTR(pool+2)); /* deep copy, not borrowed BSTR */
        VariantClear(stack); V_VT(stack) = VT_BOOL; V_BOOL(stack) = 0;
        ++f.pc;
        assert(vm(&f,0,0,0) == 0 && f.pc == 5 && f.sp == 0);
    }
    {
        Instruction code[] = {{STMT,1,2,3},{HALT,0,0,0},{PUSH_BOOL,1,0,0},{HALT,0,0,0}};
        Frame f = frame(code,4,pool,3,stack,8);
        assert(vm(&f,0,0,0) == 1 && f.request == STMT);
        ++f.pc; f.signal = 1;
        assert(vm(&f,0,0,0) == 0 && f.signal == 0 && f.sp == 1 && V_BOOL(stack));
        VariantClear(stack);
        f = frame(code,4,pool,3,stack,8);
        assert(vm(&f,0,0,0) == 1);
        ++f.pc; f.signal = 3;
        assert(vm(&f,0,0,0) == 0 && f.signal == 3 && f.sp == 0);
    }
    {
        Instruction code[] = {{CONSTANT,3,0,0},{HALT,0,0,0}};
        Frame f = frame(code,2,pool,3,stack,8);
        assert(vm(&f,0,0,0) == -1); /* invalid constant */
        code[0].a = 0; f.capacity = 0;
        assert(vm(&f,0,0,0) == -1); /* stack overflow */
        code[0].op = BINARY; f.capacity = 8;
        assert(vm(&f,0,0,0) == -1); /* underflow */
        code[0].op = JUMP; code[0].a = 99;
        assert(vm(&f,0,0,0) == -1); /* jump beyond code */
    }
    {
        Instruction code[] = {{CONSTANT,0,0,0},{CONSTANT,1,0,0},{BINARY,3,0,0},{HALT,0,0,0}};
        V_R8(pool+1) = 0;
        Frame f = frame(code,4,pool,3,stack,8);
        assert(vm(&f,0,0,0) == 1 && f.request == BINARY && f.sp == 2); /* host error, operands intact */
        VariantClear(stack); VariantClear(stack+1);
    }
    for (int i=0;i<3;++i) VariantClear(pool+i);
    puts("Native VM control, bounds, ownership and numeric checks passed");
    return 0;
}
