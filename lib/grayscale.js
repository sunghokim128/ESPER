export function grayscalePixels(pixels) {
  for (let index = 0; index < pixels.length; index += 4) {
    const gray = Math.round(
      pixels[index] * 0.2126 +
      pixels[index + 1] * 0.7152 +
      pixels[index + 2] * 0.0722
    );
    pixels[index] = gray;
    pixels[index + 1] = gray;
    pixels[index + 2] = gray;
  }
}
