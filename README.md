# summernote-ext-paste-from-word

A [Summernote](https://summernote.org/) plugin that detects HTML pasted from Microsoft Word (desktop and Word Online) and converts it to clean, minimal HTML — preserving visual formatting while removing MSO-specific markup noise.

Also handles content pasted from **Microsoft Excel** (desktop and Excel Online).

---

## Features

- Detects and cleans content from:
  - **Word desktop** (MSO namespace, `MsoNormal`/`MsoHeading` classes, `mso-list` styles, `<o:p>` tags)
  - **Word Online** (ListContainerWrapper format, `data-listid`, `color: windowtext`)
  - **Excel desktop and Excel Online** (bakes class-based styles into inline styles, removes column/colgroup elements)
- Converts Word heading styles (`MsoHeading1`–`6`, `role="heading"`, `data-ccp-parastyle`) to proper `<h1>`–`<h6>` elements
- Reconstructs nested `<ul>`/`<ol>` lists from flat MSO list markup and Word Online list wrappers
- Normalizes table borders: verbose Word Online longhand properties (`border-width`, `border-style`, `border-color`) are collapsed into a single `border` shorthand; Excel `.5pt` borders are converted to `1px`
- Preserves empty paragraphs as `<p><br></p>` — Word uses blank paragraphs for visual spacing
- Strips noise: conditional comments, MSO classes, non-visual inline styles, empty spans, `&nbsp;` artifacts, `font-size` on structural table elements
- Keeps visual formatting: `font-weight`, `font-style`, `text-decoration`, `color`, `background-color`, `border` on table cells, `border-collapse` on tables
- Falls through silently for non-Word content — regular paste behaviour is unaffected

---

## Requirements

- jQuery
- Summernote ≥ 0.8 (lite, Bootstrap 3/4/5)

---

## Installation

Include the plugin script **after** jQuery and Summernote:

```html
<script src="jquery.min.js"></script>
<script src="summernote-lite.min.js"></script>
<script src="summernote-ext-paste-from-word.js"></script>
```

---

## Usage

The plugin activates automatically — no extra configuration required:

```js
$('#editor').summernote();
```

Word content pasted into the editor is intercepted, cleaned, and inserted as tidy HTML.

---

## Custom `onPaste` callback

If you register a custom `onPaste` callback, Summernote fires it for every paste event. The plugin stores the cleaned HTML on the native event object so your callback can use it:

```js
$('#editor').summernote({
  callbacks: {
    onPaste: function(e) {
      var nativeEvent = e.originalEvent || e;

      if (nativeEvent._pfwCleanedHtml !== undefined) {
        // Paste originated from Word/Excel — use the cleaned HTML
        e.preventDefault();
        var cleaned = nativeEvent._pfwCleanedHtml;
        $(this).summernote('pasteHTML', cleaned);
      } else {
        // Regular paste — handle normally
        e.preventDefault();
        var html = nativeEvent.clipboardData.getData('text/html');
        $(this).summernote('pasteHTML', html);
      }
    }
  }
});
```

**Note:** `_pfwCleanedHtml` is set on the native event only when Word/Excel content is detected. For all other paste events the property is absent.

---

## What gets removed

| Removed | Kept |
|---------|------|
| MSO namespace markup (`xmlns:o`, `<o:p>`) | `<b>`, `<i>`, `<u>`, `<s>` |
| `MsoNormal`, `MsoList*`, `MsoHeading*` classes | `<a>` with `href` |
| `mso-*` inline styles | `font-weight`, `font-style`, `text-decoration` |
| Conditional comments (`<!--[if …]>`) | `color`, `background-color` |
| Word Online wrapper divs (`SCXW*`, `BCX*`) | `<h1>`–`<h6>` (converted from headings) |
| `color: windowtext` | `<ul>`, `<ol>`, `<li>` (reconstructed) |
| Empty `<span>` and `<p>` elements | `<table>`, `<td>` (Excel) |
| Superfluous `&nbsp;` padding | |

---

## License

MIT — see [LICENSE](LICENSE)
