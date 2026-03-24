static void CheckParity(int[] bits)
{
    if (bits == null || bits.Length != 32 || !bits.All(b => b == 0 || b == 1))
    {
        Console.WriteLine("Invalid input. Must be exactly 32 bits (0s and 1s).");
        return;
    }

    int xorResult = 0;
    for (int i = 0; i < 8; i++)
    {
        int nibble = 0;
        for (int j = 0; j < 4; j++)
            nibble = (nibble << 1) | bits[i * 4 + j];

        xorResult ^= nibble;
    }

    if (xorResult == 0)
        Console.WriteLine("✅ Parity Check PASSED");
    else
        Console.WriteLine($"❌ Parity Check FAILED — XOR result: {Convert.ToString(xorResult, 2).PadLeft(4, '0')}");
}
