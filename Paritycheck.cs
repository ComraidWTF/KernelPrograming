static void CheckParity(string bits)
{
    if (bits == null || bits.Length != 32 || !bits.All(c => c == '0' || c == '1'))
    {
        Console.WriteLine("Invalid input. Must be exactly 32 bits (0s and 1s).");
        return;
    }

    int xorResult = 0;
    for (int i = 0; i < 8; i++)
    {
        xorResult ^= Convert.ToInt32(bits.Substring(i * 4, 4), 2);
    }

    if (xorResult == 0)
        Console.WriteLine("✅ Parity Check PASSED");
    else
        Console.WriteLine($"❌ Parity Check FAILED — XOR result: {Convert.ToString(xorResult, 2).PadLeft(4, '0')}");
}
