export const EMAIL_RE =
  /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;

export const PHONE_RE =
  /(?:\+?1[\s.-]?)?(?:\(?\d{3}\)?[\s.-]?)\d{3}[\s.-]?\d{4}/g;

export const SSN_RE =
  /\b\d{3}-\d{2}-\d{4}\b/g;

export const CARD_CANDIDATE_RE =
  /\b(?:\d[ -]*?){13,19}\b/g;
