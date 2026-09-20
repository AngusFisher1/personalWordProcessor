import type { StyleId } from './model';

/**
 * Work out what a paragraph *is* when the document does not say.
 *
 * Real documents are mostly directly formatted rather than styled. Across a
 * sample of 44 of them, 1,348 of 1,573 paragraphs carried no w:pStyle at all
 * and only 11 used Heading1 - so a mapping table keyed on style names is
 * correct and almost never fires, and every heading lands on Body.
 *
 * This reads the formatting instead: Word's own outline level first, then the
 * signals a person actually uses to make something look like a heading - a
 * rule underneath it, larger text, bold, capitals - gated on the paragraph
 * being short, because a long paragraph is never a heading however it is set.
 *
 * Pure and dependency-free so it can be reasoned about and tested on its own.
 */

export interface ParaSignals {
  /** A w:pStyle we recognized. When set, nothing here is inferred. */
  explicit: StyleId | null;
  /** w:outlineLvl, 0-based. Word's own statement that this is a heading. */
  outlineLvl: number | null;
  /**
   * The explicit run size carrying the most characters, in half-points, or
   * null to inherit. The dominant size rather than the largest: one small
   * superscript should not describe the whole paragraph.
   */
  sizeHalfPt: number | null;
  /**
   * Share of characters that are bold, 0 to 1. A share rather than a flag,
   * because headings are routinely only partly bold - "Job Title, Company"
   * with the title emphasized is the common resume idiom.
   */
  boldShare: number;
  /** w:caps, or the text is written in capitals. */
  caps: boolean;
  /** The paragraph has a bottom border - the resume section-rule idiom. */
  ruled: boolean;
  centered: boolean;
  /** Rendered character count. */
  textLen: number;
  /** Carries w:numPr, so it is a list item whatever else it looks like. */
  isList: boolean;
  /** Contains an email, or a phone-length run of digits. */
  contactish: boolean;
}

/** Longer than this and it is prose, not a heading. */
const HEADING_MAX_CHARS = 80;
/**
 * Larger than the body text at all. A ratio was too blunt: a heading is
 * routinely only one size step up - 11pt in a 10pt document - and any
 * threshold loose enough to catch that is indistinguishable from "larger".
 * Shortness is what keeps this from firing on prose.
 */
/** Large enough to be the document's title rather than a section of it. */
const HUGE = 1.25;

/**
 * The body text size, as a weighted mode: whichever size carries the most
 * characters wins. Self-calibrating, so a document set in 10pt and one set in
 * 12pt are both measured against themselves rather than against a constant.
 */
export function bodySizeOf(
  signals: ParaSignals[],
  docDefaultHalfPt: number | null
): number {
  const bySize = new Map<number, number>();
  let inherited = 0;
  for (const s of signals) {
    if (s.textLen === 0) continue;
    if (s.sizeHalfPt === null) inherited += s.textLen;
    else bySize.set(s.sizeHalfPt, (bySize.get(s.sizeHalfPt) ?? 0) + s.textLen);
  }
  let best = docDefaultHalfPt ?? 22; // 11pt
  let bestWeight = inherited;
  for (const [size, weight] of bySize) {
    if (weight > bestWeight) {
      best = size;
      bestWeight = weight;
    }
  }
  return best;
}

/**
 * Half the line being bold is enough for boldness to be its point. The
 * common resume idiom - "Staff Engineer, Northwind" bold followed by an
 * unbolded date range - lands just under 60%, so a stricter threshold misses
 * exactly the case this exists for.
 */
const MOSTLY_BOLD = 0.5;

function classify(s: ParaSignals, bodySize: number): StyleId {
  // Word said so.
  if (s.outlineLvl !== null) {
    return s.outlineLvl <= 1 ? 'SectionHeading' : 'JobTitle';
  }
  if (s.textLen === 0 || s.textLen > HEADING_MAX_CHARS) return 'Body';

  const size = s.sizeHalfPt ?? bodySize;
  const big = size > bodySize;
  const huge = size >= bodySize * HUGE;
  const bold = s.boldShare >= MOSTLY_BOLD;

  if (s.ruled) return 'SectionHeading';
  if (huge) return 'SectionHeading'; // may be promoted to Name below
  if (s.caps && (bold || big)) return 'SectionHeading';
  if (big) return 'SectionHeading';
  if (s.caps && s.textLen <= 40) return 'SectionHeading';
  if (bold) return 'JobTitle';
  return 'Body';
}

/**
 * Classify every paragraph, then promote the document's title.
 *
 * The title pass is separate because it is a statement about the document,
 * not about the paragraph: the first short line that is the largest thing in
 * the document is its title, whether that is a person's name or a contract's.
 */
export function inferStyles(
  signals: ParaSignals[],
  docDefaultHalfPt: number | null
): StyleId[] {
  const bodySize = bodySizeOf(signals, docDefaultHalfPt);
  const out = signals.map((s) => {
    if (s.isList) return 'Bullet' as StyleId;
    return s.explicit ?? classify(s, bodySize);
  });

  // The largest size anywhere, so "biggest thing in the document" is a fact
  // rather than a guess.
  let maxSize = bodySize;
  for (const s of signals) {
    if (s.textLen > 0 && s.sizeHalfPt !== null && s.sizeHalfPt > maxSize) {
      maxSize = s.sizeHalfPt;
    }
  }

  const head = Math.min(signals.length, 6);
  for (let i = 0; i < head; i++) {
    const s = signals[i];
    if (s.textLen === 0 || s.isList || s.explicit) continue;
    const size = s.sizeHalfPt ?? bodySize;
    const isTitle = size === maxSize && size > bodySize && s.textLen <= 60;
    if (!isTitle) continue;
    out[i] = 'Name';

    // The line under a title is a contact line only when it looks like one.
    // Without that check a contract's date line would be set in contact type.
    const next = signals[i + 1];
    if (
      next &&
      !next.isList &&
      !next.explicit &&
      next.textLen > 0 &&
      next.textLen <= 120 &&
      (next.sizeHalfPt ?? bodySize) <= bodySize &&
      next.contactish
    ) {
      out[i + 1] = 'Contact';
    }
    break;
  }

  return out;
}

/** An email address, or a run of digits long enough to be a phone number. */
export function looksLikeContact(text: string): boolean {
  if (/[^\s@]+@[^\s@]+\.[^\s@]+/.test(text)) return true;
  const digits = text.replace(/\D/g, '');
  return digits.length >= 7 && /[()\-.\s]\d/.test(text);
}
