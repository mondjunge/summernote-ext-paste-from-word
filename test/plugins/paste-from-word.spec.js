/**
 * paste-from-word.spec.js
 * Tests for the WordCleaner class (paste-from-word plugin logic).
 */

import { describe, it, expect, beforeEach } from 'vitest';
import WordCleaner from '@/js/plugin/paste-from-word';

describe('paste-from-word: detection', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('detects Word content via xmlns:o namespace', () => {
    expect(cleaner.isWordContent(
      '<html xmlns:o="urn:schemas-microsoft-com:office:office"><body><p>text</p></body></html>'
    )).toBe(true);
  });

  it('detects Word content via ProgId', () => {
    expect(cleaner.isWordContent('<!-- ProgId=Word.Document --><p>text</p>')).toBe(true);
  });

  it('detects Word content via Mso class prefix', () => {
    expect(cleaner.isWordContent('<p class="MsoNormal">text</p>')).toBe(true);
  });

  it('detects Word content via o:p tag', () => {
    expect(cleaner.isWordContent('<p>text<o:p></o:p></p>')).toBe(true);
  });

  it('detects Word content via mso-list style', () => {
    expect(cleaner.isWordContent('<p style="mso-list:l0 level1 lfo1">item</p>')).toBe(true);
  });

  it('detects Word Online full-document content via color: windowtext', () => {
    expect(cleaner.isWordContent('<p style="color: windowtext">text</p>')).toBe(true);
  });

  it('detects Word Online content via border-bottom: 1px solid transparent', () => {
    expect(cleaner.isWordContent('<span style="border-bottom: 1px solid transparent">text</span>')).toBe(true);
  });

  it('does not flag plain HTML as Word content', () => {
    expect(cleaner.isWordContent('<p>Hello <strong>world</strong></p>')).toBe(false);
  });
});

describe('paste-from-word: conditional comment removal', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('removes [if !supportLists] blocks entirely including inner content', () => {
    const html = '<!--[if !supportLists]><span>·&nbsp;&nbsp;</span><![endif]-->Item';
    const result = cleaner.removeConditionalComments(html);
    expect(result).not.toContain('[if');
    expect(result).not.toContain('·');
    expect(result).toContain('Item');
  });

  it('strips other conditional comment markers but keeps inner content', () => {
    const html = '<!--[if gte mso 9]><xml>ignored</xml><![endif]-->';
    const result = cleaner.removeConditionalComments(html);
    expect(result).not.toContain('[if');
    expect(result).not.toContain('[endif]');
  });

  it('removes XML processing instructions', () => {
    const html = '<?xml version="1.0"?><p>text</p>';
    const result = cleaner.removeConditionalComments(html);
    expect(result).not.toContain('<?xml');
    expect(result).toContain('<p>text</p>');
  });
});

describe('paste-from-word: body extraction', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('extracts content between <body> tags', () => {
    const html = '<html><head><style>p{}</style></head><body><p>content</p></body></html>';
    expect(cleaner.extractBodyContent(html)).toBe('<p>content</p>');
  });

  it('returns the full string unchanged when no body tag is present', () => {
    const html = '<p>standalone</p>';
    expect(cleaner.extractBodyContent(html)).toBe('<p>standalone</p>');
  });
});

describe('paste-from-word: heading conversion', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('converts MsoHeading1 paragraph to h1', () => {
    const result = cleaner.clean('<p class="MsoHeading1">Introduction</p>');
    expect(result).toContain('<h1>');
    expect(result).toContain('Introduction');
    expect(result).not.toContain('<p');
  });

  [2, 3, 4, 5, 6].forEach(n => {
    it(`converts MsoHeading${n} to h${n}`, () => {
      const result = cleaner.clean(`<p class="MsoHeading${n}">Title</p>`);
      expect(result).toContain(`<h${n}>`);
    });
  });

  it('preserves inline formatting inside headings', () => {
    const result = cleaner.clean('<p class="MsoHeading1"><strong>Bold</strong> heading</p>');
    expect(result).toContain('<h1>');
    expect(result).toContain('<strong>Bold</strong>');
  });

  it('leaves non-heading paragraphs unchanged', () => {
    const result = cleaner.clean('<p class="MsoNormal">Normal text</p>');
    expect(result).not.toContain('<h1>');
  });

  it('converts Word Online role=heading aria-level=1 to h1', () => {
    const result = cleaner.clean('<p role="heading" aria-level="1">Heading One</p>');
    expect(result).toContain('<h1>');
    expect(result).toContain('Heading One');
    expect(result).not.toContain('<p');
  });

  [2, 3, 4, 5, 6].forEach(n => {
    it(`converts Word Online role=heading aria-level=${n} to h${n}`, () => {
      const result = cleaner.clean(`<p role="heading" aria-level="${n}">Title</p>`);
      expect(result).toContain(`<h${n}>`);
    });
  });

  it('converts Word Online heading with span content correctly', () => {
    const html = '<p class="Paragraph" role="heading" aria-level="1">' +
      '<span class="TextRun"><span class="NormalTextRun" data-ccp-parastyle="heading 1">My Heading</span></span>' +
      '<span class="EOP">&nbsp;</span></p>';
    const result = cleaner.clean(html);
    expect(result).toContain('<h1>');
    expect(result).toContain('My Heading');
  });

  [1, 2, 3, 4, 5, 6].forEach(n => {
    it(`converts data-ccp-parastyle="heading ${n}" (no aria-level) to h${n}`, () => {
      const html = `<p><span style="font-size: 14pt;"><span data-ccp-parastyle="heading ${n}">Title</span></span></p>`;
      const result = cleaner.clean(html);
      expect(result).toContain(`<h${n}>`);
    });
  });

  it('converts custom heading style (data-ccp-parastyle="heading 20") via font-size heuristic', () => {
    const html = '<p style="font-weight: bold; color: rgb(46, 117, 182);">' +
      '<span style="font-size: 14pt; font-weight: bold;" data-contrast="none">' +
      '<span data-ccp-parastyle="heading 20">Custom Heading</span></span>' +
      '<span class="EOP">&nbsp;</span></p>';
    const result = cleaner.clean(html);
    expect(result).toContain('<h3>');
    expect(result).toContain('Custom Heading');
  });

  it('maps 20pt+ custom heading to h1', () => {
    const html = '<p><span style="font-size: 20pt;"><span data-ccp-parastyle="heading 15">Big</span></span></p>';
    expect(cleaner.clean(html)).toContain('<h1>');
  });

  it('maps 16pt custom heading to h2', () => {
    const html = '<p><span style="font-size: 16pt;"><span data-ccp-parastyle="heading 10">Med</span></span></p>';
    expect(cleaner.clean(html)).toContain('<h2>');
  });
});

describe('paste-from-word: deduplicate inherited styles', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('removes span styles already set on parent p', () => {
    // Both p and span have same color and font-weight; span additionally has font-size
    const html = '<p style="color: rgb(46, 117, 182); font-weight: bold;">' +
      '<span style="color: rgb(46, 117, 182); font-weight: bold; font-size: 14pt;">Text</span></p>';
    const result = cleaner.clean(html);
    // span should only retain font-size (the non-redundant property)
    expect(result).toContain('font-size: 14pt');
    expect(result).not.toMatch(/span[^>]*color/);
    expect(result).not.toMatch(/span[^>]*font-weight/);
  });

  it('unwraps span entirely when all its styles duplicate parent', () => {
    const html = '<p style="color: rgb(46, 117, 182); font-weight: bold;">' +
      '<span style="color: rgb(46, 117, 182); font-weight: bold;">Text</span></p>';
    const result = cleaner.clean(html);
    expect(result).not.toContain('<span');
    expect(result).toContain('Text');
  });

  it('preserves span styles that differ from parent', () => {
    const html = '<p style="color: rgb(0, 0, 0);">' +
      '<span style="color: rgb(46, 117, 182); font-weight: bold;">Text</span></p>';
    const result = cleaner.clean(html);
    // span color differs from parent, must be kept
    expect(result).toContain('color: rgb(46, 117, 182)');
  });
});

describe('paste-from-word: list conversion', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('converts MsoListBullet paragraphs to a ul', () => {
    const html = [
      '<p class="MsoListBullet" style="mso-list:l0 level1 lfo1">',
      '<span style="mso-list:Ignore">·</span>First item',
      '</p>',
      '<p class="MsoListBullet" style="mso-list:l0 level1 lfo1">',
      '<span style="mso-list:Ignore">·</span>Second item',
      '</p>',
    ].join('');

    const result = cleaner.clean(html);
    expect(result).toContain('<ul>');
    expect(result).toContain('<li>');
    expect(result).toContain('First item');
    expect(result).toContain('Second item');
    expect(result).not.toContain('class=');
    expect(result).not.toContain('mso-list');
  });

  it('converts MsoListNumber paragraphs to an ol', () => {
    const html = [
      '<p class="MsoListNumber" style="mso-list:l0 level1 lfo1">',
      '<span style="mso-list:Ignore">1.</span>First',
      '</p>',
      '<p class="MsoListNumber" style="mso-list:l0 level1 lfo1">',
      '<span style="mso-list:Ignore">2.</span>Second',
      '</p>',
    ].join('');

    const result = cleaner.clean(html);
    expect(result).toContain('<ol>');
    expect(result).toContain('<li>');
  });

  it('builds nested ul based on level attribute', () => {
    const html = [
      '<p class="MsoListBullet" style="mso-list:l0 level1 lfo1">',
      '<span style="mso-list:Ignore">·</span>Top level',
      '</p>',
      '<p class="MsoListBullet" style="mso-list:l0 level2 lfo1">',
      '<span style="mso-list:Ignore">o</span>Nested',
      '</p>',
    ].join('');

    const result = cleaner.clean(html);
    expect(result).toMatch(/<li>[\s\S]*<ul>[\s\S]*Nested[\s\S]*<\/ul>[\s\S]*<\/li>/);
  });

  it('strips the mso-list:Ignore span from list item content', () => {
    const html =
      '<p class="MsoListBullet" style="mso-list:l0 level1 lfo1">' +
      '<span style="mso-list:Ignore">·&nbsp;&nbsp;</span>Item text' +
      '</p>';

    const result = cleaner.clean(html);
    const li = result.match(/<li>([\s\S]*?)<\/li>/);
    expect(li).not.toBeNull();
    expect(li[1]).not.toContain('·');
    expect(li[1]).toContain('Item text');
  });
});

describe('paste-from-word: style cleaning', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('removes mso-* style properties', () => {
    const result = cleaner.clean(
      '<p style="mso-margin-top-alt:auto;mso-margin-bottom-alt:auto">text</p>'
    );
    expect(result).not.toContain('mso-');
  });

  it('preserves color, font-size, font-weight, font-style, text-decoration', () => {
    const result = cleaner.clean(
      '<span style="color:red;font-size:14pt;font-weight:bold;font-style:italic;text-decoration:underline;mso-bidi-font-weight:normal">text</span>'
    );
    expect(result).toContain('color:red');
    expect(result).toContain('font-size:14pt');
    expect(result).toContain('font-weight:bold');
    expect(result).toContain('font-style:italic');
    expect(result).toContain('text-decoration:underline');
    expect(result).not.toContain('mso-bidi-font-weight');
  });

  it('removes font-family and margin properties', () => {
    const result = cleaner.clean(
      '<p style="font-family:Calibri,sans-serif;margin-left:36pt">text</p>'
    );
    expect(result).not.toContain('font-family');
    expect(result).not.toContain('margin-left');
  });

  it('removes the style attribute when no visual properties remain', () => {
    const result = cleaner.clean(
      '<p style="mso-style-unhide:no;mso-style-qformat:yes">text</p>'
    );
    expect(result).not.toContain('style=');
  });
});

describe('paste-from-word: attribute cleaning', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('removes class attributes', () => {
    const result = cleaner.clean('<p class="MsoNormal">text</p>');
    expect(result).not.toContain('class=');
  });

  it('removes lang attributes', () => {
    const result = cleaner.clean('<p lang="de-DE">text</p>');
    expect(result).not.toContain('lang=');
  });

  it('preserves href on links', () => {
    const result = cleaner.clean('<a href="https://example.com" class="MsoHyperlink">link</a>');
    expect(result).toContain('href="https://example.com"');
    expect(result).not.toContain('class=');
  });

  it('preserves colspan and rowspan on table cells', () => {
    const result = cleaner.clean(
      '<table><tr><td colspan="2" class="MsoTableGrid">cell</td></tr></table>'
    );
    expect(result).toContain('colspan="2"');
    expect(result).not.toContain('class=');
  });
});

describe('paste-from-word: noise removal', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('removes empty o:p tags', () => {
    const result = cleaner.clean('<p>text<o:p></o:p></p>');
    expect(result).not.toContain('o:p');
  });

  it('removes images with relative (Word-local) src', () => {
    const result = cleaner.clean('<p><img src="image001.png" alt="logo"></p>');
    expect(result).not.toContain('<img');
  });

  it('keeps images with absolute https src', () => {
    const result = cleaner.clean('<p><img src="https://example.com/logo.png" alt="logo"></p>');
    expect(result).toContain('<img');
    expect(result).toContain('https://example.com/logo.png');
  });

  it('keeps images with data URI src', () => {
    const result = cleaner.clean('<p><img src="data:image/png;base64,abc123" alt="x"></p>');
    expect(result).toContain('<img');
    expect(result).toContain('data:image/png');
  });
});

describe('paste-from-word: Word Online detection', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('detects Word Online content via ListContainerWrapper class', () => {
    expect(cleaner.isWordContent(
      '<div class="ListContainerWrapper SCXW123 BCX8"><ul><li>item</li></ul></div>'
    )).toBe(true);
  });

  it('detects Word Online content via data-listid attribute', () => {
    expect(cleaner.isWordContent(
      '<li data-listid="2" data-aria-level="1">item</li>'
    )).toBe(true);
  });

  it('does not flag plain HTML with div as Word Online content', () => {
    expect(cleaner.isWordContent('<div><p>Hello</p></div>')).toBe(false);
  });
});

describe('paste-from-word: Word Online list conversion', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  function makeWordOnlineItem(text, level = 1, tag = 'ul') {
    return (
      `<div class="ListContainerWrapper SCXW1 BCX8">` +
      `<${tag} class="BulletListStyle1 SCXW1 BCX8">` +
      `<li data-aria-level="${level}" data-listid="2" class="OutlineElement SCXW1 BCX8">` +
      `<p class="Paragraph SCXW1 BCX8">` +
      `<span class="TextRun SCXW1 BCX8"><span class="NormalTextRun SCXW1 BCX8">${text}</span></span>` +
      `<span class="EOP SCXW1 BCX8">&nbsp;</span>` +
      `</p></li></${tag}>` +
      `</div>`
    );
  }

  it('converts Word Online bullet list to ul', () => {
    const html = makeWordOnlineItem('First') + makeWordOnlineItem('Second');
    const result = cleaner.clean(html);
    expect(result).toContain('<ul>');
    expect(result).toContain('<li>');
    expect(result).toContain('First');
    expect(result).toContain('Second');
    expect(result).not.toContain('ListContainerWrapper');
    expect(result).not.toContain('class=');
  });

  it('converts Word Online numbered list to ol', () => {
    const html = makeWordOnlineItem('First', 1, 'ol') + makeWordOnlineItem('Second', 1, 'ol');
    const result = cleaner.clean(html);
    expect(result).toContain('<ol>');
    expect(result).toContain('<li>');
  });

  it('builds nested list based on data-aria-level', () => {
    const html = makeWordOnlineItem('Top') + makeWordOnlineItem('Nested', 2);
    const result = cleaner.clean(html);
    expect(result).toMatch(/<li>[\s\S]*<ul>[\s\S]*Nested[\s\S]*<\/ul>[\s\S]*<\/li>/);
  });

  it('removes EOP spans from list item content', () => {
    const html = makeWordOnlineItem('Item text');
    const result = cleaner.clean(html);
    const li = result.match(/<li>([\s\S]*?)<\/li>/);
    expect(li).not.toBeNull();
    expect(li[1]).not.toContain('EOP');
    expect(li[1]).toContain('Item text');
  });

  it('does not produce trailing &nbsp; in list items', () => {
    const html = makeWordOnlineItem('Clean text');
    const result = cleaner.clean(html);
    const li = result.match(/<li>([\s\S]*?)<\/li>/);
    expect(li[1].trim()).not.toMatch(/&nbsp;$/);
  });

  it('removes data-* and class attributes from converted list', () => {
    const html = makeWordOnlineItem('Item');
    const result = cleaner.clean(html);
    expect(result).not.toContain('data-aria-level');
    expect(result).not.toContain('data-listid');
    expect(result).not.toContain('class=');
  });
});

describe('paste-from-word: default style removal', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('removes color: windowtext as a default', () => {
    const result = cleaner.clean('<p style="color: windowtext">text</p>');
    expect(result).not.toContain('windowtext');
    expect(result).not.toContain('color:');
  });

  it('removes color: #000000 as a default', () => {
    const result = cleaner.clean('<p style="color: #000000">text</p>');
    expect(result).not.toContain('#000000');
  });

  it('removes background-color: transparent as a default', () => {
    const result = cleaner.clean('<p style="background-color: transparent">text</p>');
    expect(result).not.toContain('background-color');
  });

  it('removes background-color: #ffffff as a default', () => {
    const result = cleaner.clean('<p style="background-color: #ffffff">text</p>');
    expect(result).not.toContain('background-color');
  });

  it('removes font-weight: normal as a default', () => {
    const result = cleaner.clean('<p style="font-weight: normal">text</p>');
    expect(result).not.toContain('font-weight');
  });

  it('removes font-style: normal as a default', () => {
    const result = cleaner.clean('<p style="font-style: normal">text</p>');
    expect(result).not.toContain('font-style');
  });

  it('removes vertical-align: baseline as a default', () => {
    const result = cleaner.clean('<p style="vertical-align: baseline">text</p>');
    expect(result).not.toContain('vertical-align');
  });

  it('removes text-align: left as a default', () => {
    const result = cleaner.clean('<p style="text-align: left">text</p>');
    expect(result).not.toContain('text-align');
  });

  it('removes text-align: start as a default', () => {
    const result = cleaner.clean('<p style="color: windowtext; text-align: start">text</p>');
    expect(result).not.toContain('text-align');
  });

  it('keeps non-default color values', () => {
    const result = cleaner.clean('<span style="color: #2e75b6">text</span>');
    expect(result).toContain('color: #2e75b6');
  });

  it('keeps non-default background-color values', () => {
    const result = cleaner.clean(
      '<table style="color: windowtext"><tr>' +
      '<td style="background-color: #2e75b6">cell</td>' +
      '</tr></table>'
    );
    expect(result).toContain('background-color: #2e75b6');
  });

  it('keeps font-weight: bold', () => {
    const result = cleaner.clean('<span style="font-weight: bold; color: windowtext">text</span>');
    expect(result).toContain('font-weight: bold');
  });

  it('keeps text-align: center', () => {
    const result = cleaner.clean('<p style="text-align: center; color: windowtext">text</p>');
    expect(result).toContain('text-align: center');
  });

  it('strips !important from style values when comparing to defaults', () => {
    const result = cleaner.clean('<p style="color: windowtext !important">text</p>');
    expect(result).not.toContain('windowtext');
  });
});

describe('paste-from-word: div unwrapping and list merging', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('unwraps wrapper divs around content', () => {
    const result = cleaner.clean('<p style="color: windowtext"><div><p>text</p></div></p>');
    // After unwrapping: just text paragraph
    expect(result).toContain('text');
    expect(result).not.toContain('<div');
  });

  it('merges consecutive ul elements into one', () => {
    const html =
      '<p style="color: windowtext"><ul><li>First</li></ul><ul><li>Second</li></ul></p>';
    // Use windowtext to trigger detection
    const result = cleaner.clean(
      '<div style="color: windowtext">' +
      '<div><ul><li>First</li></ul></div>' +
      '<div><ul><li>Second</li></ul></div>' +
      '</div>'
    );
    // Should produce a single ul with two items
    const uls = result.match(/<ul>/g);
    expect(uls).toHaveLength(1);
    expect(result).toContain('First');
    expect(result).toContain('Second');
  });

  it('merges consecutive ol elements into one', () => {
    const result = cleaner.clean(
      '<div style="color: windowtext">' +
      '<div><ol><li>Step 1</li></ol></div>' +
      '<div><ol><li>Step 2</li></ol></div>' +
      '</div>'
    );
    const ols = result.match(/<ol>/g);
    expect(ols).toHaveLength(1);
  });

  it('does not merge ul and ol together', () => {
    const result = cleaner.clean(
      '<div style="color: windowtext">' +
      '<div><ul><li>Bullet</li></ul></div>' +
      '<div><ol><li>Number</li></ol></div>' +
      '</div>'
    );
    expect(result).toContain('<ul>');
    expect(result).toContain('<ol>');
  });

  it('unwraps <p> inside <li> elements', () => {
    const result = cleaner.clean(
      '<div style="color: windowtext">' +
      '<ul><li><p>Item text</p></li></ul>' +
      '</div>'
    );
    expect(result).toContain('<li>Item text</li>');
    expect(result).not.toContain('<li><p>');
  });

  it('unwraps <p> inside table cells', () => {
    const result = cleaner.clean(
      '<div style="color: windowtext">' +
      '<table><tr><td><p>Cell text</p></td></tr></table>' +
      '</div>'
    );
    expect(result).toContain('<td>Cell text</td>');
    expect(result).not.toContain('<td><p>');
  });

  it('separates multiple <p> in a <td> with <br>', () => {
    const result = cleaner.clean(
      '<div style="color: windowtext">' +
      '<table><tr><td><p>Line 1</p><p>Line 2</p></td></tr></table>' +
      '</div>'
    );
    expect(result).toContain('Line 1');
    expect(result).toContain('Line 2');
    expect(result).toContain('<br>');
  });
});

describe('paste-from-word: empty block removal', () => {
  let cleaner;
  beforeEach(() => { cleaner = new WordCleaner(); });

  it('removes paragraphs left empty after cleanup', () => {
    const result = cleaner.clean('<p class="MsoNormal"><o:p>&nbsp;</o:p></p><p>real content</p>');
    expect(result).toContain('real content');
  });

  it('keeps elements that contain only a br', () => {
    const result = cleaner.clean('<p><br></p>');
    expect(result).toContain('<br>');
  });
});
