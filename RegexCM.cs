var regex = new Regex(
    @"^(CM\d+):\s*(.*?)(?=^CM\d+:|\z)",
    RegexOptions.Multiline | RegexOptions.Singleline
);
