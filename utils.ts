export function safeIsSupported(isSupported: () => boolean): boolean {
  try {
    const supported = isSupported();
    return supported;
  } catch {
    return false;
  }
}
