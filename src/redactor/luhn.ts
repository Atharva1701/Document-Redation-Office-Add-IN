export function isLikelyCreditCard(raw: string): boolean {
  const digits = raw.replace(/[^\d]/g, "");
  if (digits.length < 13 || digits.length > 19) return false;

  // Luhn check
  let sum = 0;
  let doubleIt = false;

  for (let i = digits.length - 1; i >= 0; i--) {
    let n = digits.charCodeAt(i) - 48;
    if (doubleIt) {
      n *= 2;
      if (n > 9) n -= 9;
    }
    sum += n;
    doubleIt = !doubleIt;
  }
  return sum % 10 === 0;
}
