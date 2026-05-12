# Product

## Register

product

## Users

Two equally important audiences:

1. **P6 Schedulers / Project Controls Engineers** — power users who live in Primavera P6. They need exact variance numbers, task codes, relationship changes, and data-date awareness. They trust numbers over narrative.
2. **Project Managers** — non-technical stakeholders who need to understand schedule health quickly. They read the AI summary, look at the charts, and want a clear answer: are we on track or not?

Both users review the same output, often in the same meeting. The interface must serve precision-seekers and summary-readers simultaneously.

## Product Purpose

A web tool that accepts two Primavera P6 XER files (baseline and updated schedule), compares them, and produces an interactive dashboard with variance data, charts, and an AI-written narrative. Replaces hours of manual schedule diffing in Excel.

Success: a scheduler uploads two XER files and within 10 seconds has a shareable, printable view of exactly what changed and why it matters.

## Brand Personality

Authoritative, precise, clear.

The tool should feel like a trusted senior colleague — professional without being cold, confident without being flashy. It earns trust through information density done right, not through decoration.

## Anti-references

- **MS Project / Primavera P6 UI** — grey, cluttered, legacy enterprise. Exactly what users are trying to escape.
- **Generic Bootstrap dashboards** — blue primary buttons, DataTables default styling, cookie-cutter sidebar layouts.
- **Overly minimal / startup-white** — this is a data-heavy tool; it cannot sacrifice density for aesthetic minimalism.

## Design Principles

1. **Data is the hero** — every visual decision should make the numbers easier to read, not harder. No decoration that competes with content.
2. **Trust through density done right** — pack information in, but use hierarchy and whitespace to make it scannable, not overwhelming.
3. **One truth per screen** — the dashboard answers one question: what changed and does it matter? Don't distract from that.
4. **Earn confidence immediately** — the first thing a user sees after uploading should feel like a credible, professional output they can share with a client.
5. **Both audiences, same screen** — the AI summary card is for PMs, the variance table is for schedulers. Neither should dominate.

## Accessibility & Inclusion

- WCAG AA minimum (contrast ratios, keyboard navigation)
- Color should never be the only signal (e.g. red/green variance also needs +/- sign or icon)
- Table data must be readable without charts (for printed reports and screen readers)
