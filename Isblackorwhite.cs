var (r, g, b) = (pixel.R, pixel.G, pixel.B);

bool black = r <= tolerance && g <= tolerance && b <= tolerance;
bool white = r >= 255 - tolerance && g >= 255 - tolerance && b >= 255 - tolerance;
