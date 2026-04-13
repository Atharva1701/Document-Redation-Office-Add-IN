import { EMAIL_RE, PHONE_RE, SSN_RE, CARD_CANDIDATE_RE } from "./patterns";
import { isLikelyCreditCard } from "./luhn";

export type RedactionResult = {
  emails: number;
  phones: number;
  ssns: number;
  cards: number;
  note?: string;
};

function findAll(text: string, re: RegExp): string[] {
  const safe = new RegExp(
    re.source,
    re.flags.includes("g") ? re.flags : re.flags + "g"
  );
  return Array.from(text.matchAll(safe), m => m[0]);
}

async function tryEnableTrackChanges(context: Word.RequestContext): Promise<boolean> {
  const hasWordApi15 = Office.context.requirements.isSetSupported("WordApi", "1.5");
  const hasDesktopTrack =
    Office.context.requirements.isSetSupported("WordApiDesktop", "1.4");

  if (!hasWordApi15 || !hasDesktopTrack) return false;

  (context.document as any).trackRevisions = true;
  await context.sync();
  return true;
}


async function ensureConfidentialHeader(context: Word.RequestContext): Promise<boolean> {
  const body = context.document.body;
  body.load("text");
  await context.sync();

  const header = "CONFIDENTIAL DOCUMENT";
  const currentText = (body.text || "").trimStart();

  if (currentText.startsWith(header)) return false;

  body.insertParagraph(header, Word.InsertLocation.start);
  await context.sync();
  return true;
}

async function redactByParagraph(
  context: Word.RequestContext,
  re: RegExp,
  replacement: string,
  validator?: (s: string) => boolean
): Promise<number> {
  const paras = context.document.body.paragraphs;
  paras.load("items/text");
  await context.sync();

  let count = 0;

  for (const p of paras.items) {
    const original = p.text;
    const matches = findAll(original, re)
      .filter(m => (validator ? validator(m) : true));

    if (matches.length === 0) continue;

    let updated = original;
    matches.sort((a, b) => b.length - a.length);

    for (const m of matches) {
      updated = updated.split(m).join(replacement);
      count++;
    }

    p.insertText(updated, Word.InsertLocation.replace);
  }

  await context.sync();
  return count;
}

export async function runRedaction(): Promise<RedactionResult> {
  return Word.run(async (context) => {
    const trackingEnabled = await tryEnableTrackChanges(context);
    const headerAdded = await ensureConfidentialHeader(context);

    const emailCount = await redactByParagraph(
      context,
      EMAIL_RE,
      "[EMAIL REDACTED]"
    );

    const phoneCount = await redactByParagraph(
      context,
      PHONE_RE,
      "[PHONE REDACTED]"
    );

    const ssnCount = await redactByParagraph(
      context,
      SSN_RE,
      "[SSN REDACTED]"
    );

    const cardCount = await redactByParagraph(
      context,
      CARD_CANDIDATE_RE,
      "[CREDIT CARD REDACTED]",
      isLikelyCreditCard
    );

    const noteParts: string[] = [];
    if (headerAdded) noteParts.push("Header inserted.");
    else noteParts.push("Header already present.");

    if (trackingEnabled)
      noteParts.push("Track Changes enabled.");
    else
      noteParts.push("Track Changes unavailable in this host; continued without it.");

    return {
      emails: emailCount,
      phones: phoneCount,
      ssns: ssnCount,
      cards: cardCount,
      note: noteParts.join(" ")
    };
  });
}
