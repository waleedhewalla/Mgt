# Brief: closing the pre-print verification items

You are one of several parallel reviewers closing the last open items before printing. Work only on your assigned book and range, and write only the two output files named in your task.

## Sources available in this container
Use these sources to verify. Do not rely on memory. The web is mostly blocked, but GitHub raw files work.

| Need | Command (run from /home/user/Mgt/tools/print) |
|---|---|
| Find a verse by its wording, or print a verse | `python3 qfind.py "نص من الآية"`, or `python3 qfind.py هود 9 11` |
| Hadith wording and collection number (Bukhari, Muslim, Abu Dawud, Tirmidhi, Nasa'i, Ibn Majah, Malik) | `python3 hsearch.py "عبارة من المتن" [bukhari,muslim]` |
| Tafsir of an ayah | `python3 tafsir.py <slug> <surah_no> <ayah>`. Slugs are listed in tafsir.py. Ibn Kathir, Ibn Ashur, Qurtubi, Razi, Tabari, Sa'di and al-Shanqiti (Adwa') are available; Sayyid Qutb's Zilal is not. |
| Classical books in the OpenITI corpus: lexicons, usul, Muqaddimah and others | `python3 openiti.py list <repo> <filter>`, then `get`, then grep the local file. See the docstring for the repo of each author. |
| Muqaddimah of Ibn Khaldun (full text, already downloaded) | `/tmp/claude-0/-home-user-Mgt/be2474ec-46f6-5fbc-9b19-f8201f104606/scratchpad/res/openiti/data/0808IbnKhaldun/0808IbnKhaldun.Tarikh/0808IbnKhaldun.Tarikh.Shamela0012320-ara1.mARkdown`. Headings are lines starting `### |`. Book 1 of al-'Ibar is the Muqaddimah. |
| WebSearch tool | Snippets only. Useful for modern books' publication facts. Prefer primary texts. |

## Text of the book
Each book is dumped with paragraph IDs in `/tmp/claude-0/-home-user-Mgt/be2474ec-46f6-5fbc-9b19-f8201f104606/scratchpad/w/<key>.txt`:
- `P<i>` is a body paragraph, with `[Style]` when it has a named style.
- `T<k>.<r>.<c>` is a table cell.

The keys are:
- ta = التأسيس المعرفي
- b1, b2, b3, b4 = Books One to Four of موسوعة الإدارة القرآنية
- g1, g2, g3 = the companion guides

IDs refer to the current files in /home/user/Mgt/book/. Do not edit any .docx yourself.

## What "closing" an item means
For every item, end in exactly one of four states:
1. **VERIFIED.** The source text was found. Fix the quotation, attribution, number or reference to match it exactly, and remove any draft marker.
2. **CORRECTED.** The source says something different. Rewrite the sentence so it is true to the source.
3. **REPHRASED.** The source is not reachable here, for example a modern book. Remove the quotation marks and present it as a paraphrase with the author and title (e.g. «ويرى فلان في كتابه … أن …»). Hedge only as much as needed. Never invent page numbers, editions or wording.
4. **REMOVED.** The claim cannot be supported and the argument does not need it. Delete it, or soften it so nothing false remains.

## Production markers in التأسيس المعرفي
These are not allowed in the print edition:
- ⚠, and phrases such as «من الذاكرة», «لم نقرأ», «بحسب الملخص», «أداة تحقق», «يُتحقق منه قبل النشر», and «لم يُقرأ في المطبوع».
- «ملاحظة المنهج» paragraphs that only describe the drafting process.

What to do with them:
- **⚠ marks:** after you close an item, delete its ⚠.
- **✓ marks:** a ✓ may stay only where you confirmed the text against a source here. The reading-guide table T2 defines ✓; leave the table to the coordinator.
- **«ملاحظة المنهج» paragraphs:** keep only real methodological limits, rewritten as normal scholarly notes in the book's voice. Delete pure process notes.
- **Footnote/«الهوامش» lists:** update them to the verified source (author, title, and the edition you actually checked, e.g. «نسخة المكتبة الشاملة الرقمية عبر مستودع OpenITI»).

## Edits file
Write a JSON list to your assigned path. The schema matches `apply_edits.py`:
```
{"id": "P65", "old": "exact substring occurring once in that paragraph", "new": "replacement", "why": "short reason", "state": "VERIFIED|CORRECTED|REPHRASED|REMOVED"}
{"id": "P65", "op": "delete", "why": "..."}
{"id": "P65", "op": "insert_after", "text": "new paragraph", "like": "P64", "why": "..."}
```
Before you finish, check mechanically that every `old` occurs exactly once in its paragraph's text in the dump, for example with a short Python script over the dump file.

For long rewrites (a whole paragraph), use `old` = the full paragraph text and `new` = the new text.

## Hard constraints
- Arabic: a formal, measured essay register, matching the book's own voice. Use no em dash (—) in new text; use commas, colons or brackets.
- Do not change the book's thesis or its declared degrees (قطعي/راجح/اجتهاد/فرضية) except to correct a factual error.
- Never change Quran or hadith wording except to make it match the source exactly.
- Do not add facts, quotations or references you did not verify here.
- Do not touch paragraphs outside your assigned range.

## Report
Write a short Markdown report in Arabic to your assigned path. Include:
- counts per state;
- a table of every item, with ID, what it was, the final state and the source used;
- anything you could not close, with the reason.

Your final message: a few lines with the counts and the two file paths.
