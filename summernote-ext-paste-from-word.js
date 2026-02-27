(function(factory) {
  if (typeof define === 'function' && define.amd) {
    define(['jquery'], factory);
  } else if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('jquery'));
  } else {
    factory(window.jQuery);
  }
}(function($) {
  /**
   * Copyright (c) 2026 HIS eG - Tim Wahrendorff
   * Licensed under the MIT License (http://opensource.org/licenses/MIT)
   * 
   * paste-from-word plugin for Summernote
   *
   * Detects HTML pasted from Microsoft Word (desktop and Word Online) and
   * converts it to clean HTML, preserving visual formatting while removing
   * MSO-specific and Word Online-specific markup noise.
   *
   * Usage:
   *   Include this file after summernote, then initialize normally:
   *   $('.editor').summernote({ ... });
   */
  $.extend($.summernote.plugins, {
    'paste-from-word': function(context) {
      var self = this;
      var $editable = context.layoutInfo.editable;

      // -----------------------------------------------------------------------
      // Lifecycle
      // -----------------------------------------------------------------------

      this.initialize = function() {
        self._pasteHandler = function(event) {
          var cd = event.clipboardData;
          if (!cd) return;
          var html = cd.getData('text/html');
          if (!html || !self._isWordContent(html)) return;
          console.log('[paste-from-word] Word content detected, cleaning...', html);
          var cleaned = self._cleanWordHtml(html);
          event.preventDefault();
          if (context.options.callbacks && context.options.callbacks.onPaste) {
            console.log('custom onPaste callback is registered ...');
            // A custom onPaste callback is registered — store the cleaned HTML
            // on the native event so the callback can retrieve it via
            //   (e.originalEvent || e)._pfwCleanedHtml
            // and use it instead of clipboardData.getData('text/html').
            event._pfwCleanedHtml = cleaned;
            // Let the event continue to bubble so Editor.js fires the callback.
          } else {
            // No custom paste handler — insert directly and suppress other handlers.
            event.stopImmediatePropagation();
            context.invoke('editor.pasteHTML', cleaned);
          }
        };
        $editable[0].addEventListener('paste', self._pasteHandler, true);
      };

      this.destroy = function() {
        $editable[0].removeEventListener('paste', self._pasteHandler, true);
      };

      // -----------------------------------------------------------------------
      // Detection
      // -----------------------------------------------------------------------

      /**
       * Returns true if the HTML string appears to originate from Microsoft Word
       * (desktop or Word Online).
       */
      this._isWordContent = function(html) {
        return (
          // Desktop Word (MSO)
          /xmlns:o="urn:schemas-microsoft-com/.test(html) ||
          /ProgId=Word\.Document/.test(html) ||
          /class="?Mso[A-Z]/.test(html) ||
          /<o:p[\s>]/.test(html) ||
          /mso-list\s*:/.test(html) ||
          // Word Online — ListContainerWrapper format (individual items pasted)
          /class="[^"]*ListContainerWrapper/.test(html) ||
          /data-listid=/.test(html) ||
          // Word Online — full document paste (native ul/ol with wrapper divs)
          /color:\s*windowtext/i.test(html) ||
          /border-bottom:\s*1px solid transparent/.test(html) ||
          // Excel (desktop and Online)
          self._isExcelContent(html)
        );
      };

      /**
       * Returns true if the HTML string appears to originate from Microsoft Excel
       * (desktop or Excel Online).
       */
      this._isExcelContent = function(html) {
        return (
          /content=["']?Excel\.Sheet/i.test(html) ||
          /mso-displayed-decimal-separator/.test(html) ||
          /Generator["']?\s+content=["']?Microsoft\s+Excel/i.test(html)
        );
      };

      // -----------------------------------------------------------------------
      // Main pipeline
      // -----------------------------------------------------------------------

      this._cleanWordHtml = function(html) {
        html = self._removeConditionalComments(html);
        if (self._isExcelContent(html)) html = self._preprocessExcel(html);
        html = self._extractBodyContent(html);

        var doc = new DOMParser().parseFromString(
          '<div id="__pfword__">' + html + '</div>',
          'text/html'
        );
        var container = doc.getElementById('__pfword__');
        if (!container) return html;

        self._convertHeadings(container);
        self._convertWordOnlineLists(container);
        self._convertLists(container);
        self._unwrapDivs(container);
        self._mergeSiblingLists(container);
        self._removeNoiseNodes(container);
        self._cleanStyles(container);
        self._cleanAttributes(container);
        self._cleanHeadingSpans(container);
        self._deduplicateInheritedStyles(container);
        self._unwrapEmptySpans(container);
        self._replaceNbsp(container);
        self._unwrapWhitespaceSpans(container);
        self._removeEmptyBlocks(container);

        return container.innerHTML;
      };

      // -----------------------------------------------------------------------
      // String pre-processing
      // -----------------------------------------------------------------------

      this._removeConditionalComments = function(html) {
        html = html.replace(/<!--\[if !support[^\]]*\]>[\s\S]*?<!\[endif\]-->/gi, '');
        html = html.replace(/<!--\[if[^\]]*\]>/gi, '');
        html = html.replace(/<!\[endif\]-->/gi, '');
        html = html.replace(/<\?xml[^?]*\?>/gi, '');
        return html;
      };

      this._extractBodyContent = function(html) {
        var bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
        return bodyMatch ? bodyMatch[1] : html;
      };

      /**
       * Pre-process Excel HTML before the body is extracted.
       * Excel stores visual styles in a <style> block with named classes
       * (e.g. `.xl66 { font-weight: 700 }`). This must be baked into inline
       * styles before the head is discarded, and <col>/<colgroup> elements
       * (column widths) are removed as they are not useful in rich-text context.
       */
      this._preprocessExcel = function(html) {
        var doc = new DOMParser().parseFromString(html, 'text/html');
        self._applyExcelClassStyles(doc);
        doc.querySelectorAll('col, colgroup').forEach(function(el) { el.parentNode.removeChild(el); });
        return doc.documentElement.outerHTML;
      };

      /**
       * Parse class rules from <style> blocks and apply matching visual properties
       * as inline styles on all matching elements in the document.
       * Only class selectors (`.xl66`) are processed; element rules are ignored.
       * Existing inline styles are preserved and take precedence (appended after).
       */
      this._applyExcelClassStyles = function(doc) {
        var KEEP_PROPS = ['color', 'background-color', 'font-weight', 'font-style', 'text-decoration'];

        // Collect all class rules from style blocks
        var classRules = {};
        doc.querySelectorAll('style').forEach(function(styleEl) {
          var text = styleEl.textContent || '';
          var ruleRe = /\.([a-zA-Z][\w-]*)\s*\{([^}]*)\}/g;
          var m;
          while ((m = ruleRe.exec(text)) !== null) {
            var className = m[1];
            var body = m[2];
            var declarations = [];
            body.split(';').forEach(function(decl) {
              var colon = decl.indexOf(':');
              if (colon === -1) return;
              var prop = decl.slice(0, colon).trim().toLowerCase();
              var value = decl.slice(colon + 1).trim();
              if (!value) return;
              // Normalize `background` shorthand to `background-color` (Excel uses solid colors)
              if (prop === 'background') prop = 'background-color';
              if (KEEP_PROPS.indexOf(prop) !== -1) {
                declarations.push(prop + ': ' + value);
              }
            });
            if (declarations.length) {
              classRules[className] = declarations;
            }
          }
        });

        if (!Object.keys(classRules).length) return;

        // Apply matched class rules as inline styles
        doc.querySelectorAll('[class]').forEach(function(el) {
          var classes = (el.getAttribute('class') || '').split(/\s+/);
          var toApply = [];
          classes.forEach(function(cls) {
            if (classRules[cls]) {
              classRules[cls].forEach(function(decl) { toApply.push(decl); });
            }
          });
          if (!toApply.length) return;
          var existing = el.getAttribute('style') || '';
          var combined = toApply.join('; ') + (existing ? '; ' + existing : '');
          el.setAttribute('style', combined);
        });
      };

      // -----------------------------------------------------------------------
      // DOM transformations
      // -----------------------------------------------------------------------

      this._convertHeadings = function(container) {
        var headingMap = {
          'MsoHeading1': 'h1', 'MsoHeading2': 'h2', 'MsoHeading3': 'h3',
          'MsoHeading4': 'h4', 'MsoHeading5': 'h5', 'MsoHeading6': 'h6',
        };

        container.querySelectorAll('p').forEach(function(p) {
          var headingTag = null;

          // 1. Word Online standard headings: role="heading" + aria-level (reliable W3C markers)
          if (p.getAttribute('role') === 'heading') {
            var ariaLevel = parseInt(p.getAttribute('aria-level') || '0', 10);
            if (ariaLevel >= 1 && ariaLevel <= 6) {
              headingTag = 'h' + ariaLevel;
            }
          }

          // 2. Word Online custom heading styles: data-ccp-parastyle="heading N"
          //    Built-in headings also carry this, but those are already caught above.
          //    For N <= 6 the number maps directly; for custom styles (N > 6) infer
          //    the visual level from the paragraph's font size.
          if (!headingTag) {
            var childSpan = p.querySelector('span[data-ccp-parastyle]');
            if (childSpan) {
              var ccpStyle = (childSpan.getAttribute('data-ccp-parastyle') || '').toLowerCase().trim();
              var ccpMatch = ccpStyle.match(/^heading\s+(\d+)$/);
              if (ccpMatch) {
                var n = parseInt(ccpMatch[1], 10);
                if (n >= 1 && n <= 6) {
                  headingTag = 'h' + n;
                } else {
                  headingTag = self._inferHeadingTagFromFontSize(p);
                }
              }
            }
          }

          // 3. Fallback: MSO class names (desktop Word)
          if (!headingTag) {
            var cls = p.className || '';
            for (var msoClass in headingMap) {
              if (cls.indexOf(msoClass) !== -1) {
                headingTag = headingMap[msoClass];
                break;
              }
            }
          }

          // Convert entire p element to heading if recognized
          if (headingTag) {
            var heading = container.ownerDocument.createElement(headingTag);
            heading.innerHTML = p.innerHTML;
            p.replaceWith(heading);
          }
        });
      };

      /**
       * Estimate a heading tag (h1–h5) from the largest font-size found in
       * inline styles within the paragraph. Used for custom Word heading styles
       * that have no aria-level attribute.
       */
      this._inferHeadingTagFromFontSize = function(p) {
        var maxPt = 0;
        var check = function(style) {
          var m = (style || '').match(/font-size:\s*([\d.]+)pt/i);
          if (m) { var pt = parseFloat(m[1]); if (pt > maxPt) maxPt = pt; }
        };
        check(p.getAttribute('style'));
        p.querySelectorAll('span[style]').forEach(function(s) { check(s.getAttribute('style')); });
        if (maxPt >= 20) return 'h1';
        if (maxPt >= 16) return 'h2';
        if (maxPt >= 14) return 'h3';
        if (maxPt >= 12) return 'h4';
        return 'h5';
      };

      this._convertWordOnlineLists = function(container) {
        var doc = container.ownerDocument;

        // ListContainerWrapper divs may be grandchildren (inside outer SCXW divs).
        // Collect every distinct parent node and process each one.
        var seen = [];
        var parents = [];
        container.querySelectorAll('[class*="ListContainerWrapper"]').forEach(function(el) {
          var p = el.parentNode;
          if (p && seen.indexOf(p) === -1) {
            seen.push(p);
            parents.push(p);
          }
        });

        parents.forEach(function(parent) {
          var children = Array.from(parent.childNodes);
          var i = 0;

          while (i < children.length) {
            if (!self._isWordOnlineListWrapper(children[i])) { i++; continue; }

            var group = [];
            while (i < children.length && self._isWordOnlineListWrapper(children[i])) {
              var wrapper = children[i];
              var listEl = wrapper.querySelector('ul, ol');
              var li = listEl ? listEl.querySelector('li') : null;
              if (li) {
                group.push({
                  el: wrapper,
                  level: parseInt(li.getAttribute('data-aria-level') || '1', 10),
                  isOrdered: listEl.tagName.toUpperCase() === 'OL',
                  html: self._extractWordOnlineLiContent(li),
                });
              }
              i++;
            }

            if (group.length) {
              var listRoot = self._buildNestedList(doc, group);
              group[0].el.parentNode.insertBefore(listRoot, group[0].el);
              group.forEach(function(item) {
                if (item.el.parentNode) item.el.parentNode.removeChild(item.el);
              });
            }
          }
        });
      };

      this._isWordOnlineListWrapper = function(node) {
        if (!node || node.nodeType !== 1) return false;
        return (node.getAttribute('class') || '').indexOf('ListContainerWrapper') !== -1;
      };

      this._extractWordOnlineLiContent = function(li) {
        var clone = li.cloneNode(true);
        clone.querySelectorAll('span').forEach(function(span) {
          if (/\bEOP\b/.test(span.getAttribute('class') || '')) {
            span.parentNode.removeChild(span);
          }
        });
        clone.querySelectorAll('p').forEach(function(p) {
          Array.from(p.childNodes).forEach(function(child) {
            p.parentNode.insertBefore(child, p);
          });
          p.parentNode.removeChild(p);
        });
        return clone.innerHTML.trim().replace(/\u00a0$/, '').trim();
      };

      this._convertLists = function(container) {
        var doc = container.ownerDocument;
        var children = Array.from(container.childNodes);
        var i = 0;

        while (i < children.length) {
          if (!self._isListParagraph(children[i])) { i++; continue; }

          var items = [];
          while (i < children.length && self._isListParagraph(children[i])) {
            var para = children[i];
            items.push({
              el: para,
              level: self._getListLevel(para),
              isOrdered: self._isOrderedList(para),
              html: self._extractListItemContent(para),
            });
            i++;
          }

          var listRoot = self._buildNestedList(doc, items);
          items[0].el.parentNode.insertBefore(listRoot, items[0].el);
          items.forEach(function(item) {
            if (item.el.parentNode) item.el.parentNode.removeChild(item.el);
          });
        }
      };

      this._isListParagraph = function(node) {
        if (!node || node.nodeType !== 1) return false;
        var tag = node.tagName.toUpperCase();
        if (tag !== 'P' && tag !== 'DIV') return false;
        var style = node.getAttribute('style') || '';
        var cls = node.getAttribute('class') || '';
        return /mso-list\s*:/i.test(style) || /MsoList/.test(cls);
      };

      this._getListLevel = function(para) {
        var style = para.getAttribute('style') || '';
        var match = style.match(/mso-list\s*:[^;]*level\s*(\d+)/i);
        return match ? parseInt(match[1], 10) : 1;
      };

      this._isOrderedList = function(para) {
        var cls = para.getAttribute('class') || '';
        if (/MsoListNumber/.test(cls)) return true;
        if (/MsoListBullet/.test(cls)) return false;
        var ignoreSpan = para.querySelector('[style*="mso-list:Ignore"], [style*="mso-list: Ignore"]');
        if (ignoreSpan) {
          var text = ignoreSpan.textContent.replace(/\u00a0/g, '').trim();
          if (/^[\d]+[.)]/.test(text) || /^[ivxlcdmIVXLCDM]+\./i.test(text) || /^[a-zA-Z]\./.test(text)) {
            return true;
          }
        }
        return false;
      };

      this._extractListItemContent = function(para) {
        var clone = para.cloneNode(true);
        clone.querySelectorAll('[style*="mso-list:Ignore"], [style*="mso-list: Ignore"]').forEach(function(s) {
          s.parentNode.removeChild(s);
        });
        return clone.innerHTML.trim();
      };

      this._buildNestedList = function(doc, items) {
        if (!items.length) return doc.createElement('ul');

        var rootTag = items[0].isOrdered ? 'ol' : 'ul';
        var root = doc.createElement(rootTag);
        var stack = [{ list: root, level: 1 }];

        items.forEach(function(item) {
          while (stack.length > 1 && stack[stack.length - 1].level > item.level) {
            stack.pop();
          }

          var top = stack[stack.length - 1];

          if (top.level < item.level) {
            var lastLi = top.list.lastElementChild;
            var nestedList = doc.createElement(item.isOrdered ? 'ol' : 'ul');
            (lastLi || top.list).appendChild(nestedList);
            stack.push({ list: nestedList, level: item.level });
          }

          var li = doc.createElement('li');
          li.innerHTML = item.html;
          stack[stack.length - 1].list.appendChild(li);
        });

        return root;
      };

      /**
       * Unwrap all <div> elements — Word Online wraps content in semantically
       * empty divs. Process deepest first.
       */
      this._unwrapDivs = function(container) {
        Array.from(container.querySelectorAll('div')).reverse().forEach(function(div) {
          if (div.parentNode) {
            Array.from(div.childNodes).forEach(function(child) {
              div.parentNode.insertBefore(child, div);
            });
            div.parentNode.removeChild(div);
          }
        });
      };

      /**
       * Merge consecutive <ul>/<ol> siblings of the same type.
       */
      this._mergeSiblingLists = function(container) {
        var changed = true;
        while (changed) {
          changed = false;
          container.querySelectorAll('ul + ul, ol + ol').forEach(function(list) {
            var prev = list.previousElementSibling;
            if (prev && prev.tagName === list.tagName) {
              while (list.firstChild) prev.appendChild(list.firstChild);
              list.parentNode.removeChild(list);
              changed = true;
            }
          });
        }
      };

      // -----------------------------------------------------------------------
      // Noise removal
      // -----------------------------------------------------------------------

      this._removeNoiseNodes = function(container) {
        // Remove HTML comment nodes (e.g. <!--StartFragment--> / <!--EndFragment-->
        // that Excel embeds inside table markup)
        var walker = container.ownerDocument.createTreeWalker(container, 128 /* SHOW_COMMENT */);
        var comments = [];
        var commentNode;
        while ((commentNode = walker.nextNode())) comments.push(commentNode);
        comments.forEach(function(c) { c.parentNode.removeChild(c); });

        container.querySelectorAll('o\\:p').forEach(function(el) {
          if (el.textContent.trim()) {
            el.replaceWith.apply(el, Array.from(el.childNodes));
          } else {
            el.parentNode.removeChild(el);
          }
        });

        Array.from(container.querySelectorAll('*')).forEach(function(el) {
          if (el.tagName && el.tagName.indexOf(':') !== -1 && el.parentNode) {
            Array.from(el.childNodes).forEach(function(child) {
              el.parentNode.insertBefore(child, el);
            });
            el.parentNode.removeChild(el);
          }
        });

        // Remove Word Online EOP (end-of-paragraph) spans
        container.querySelectorAll('span').forEach(function(span) {
          var cls = span.getAttribute('class') || '';
          if (/\bEOP\b/.test(cls) && span.parentNode) {
            span.parentNode.removeChild(span);
          }
        });

        // Unwrap spans carrying no visual styles
        container.querySelectorAll('span').forEach(function(span) {
          if (self._hasOnlyNoisyStyles(span) && span.parentNode) {
            Array.from(span.childNodes).forEach(function(child) {
              span.parentNode.insertBefore(child, span);
            });
            span.parentNode.removeChild(span);
          }
        });

        // Unwrap <p> inside <li> and table cells
        container.querySelectorAll('li, td, th').forEach(function(cell) {
          var doc = cell.ownerDocument;
          var paragraphs = Array.from(cell.querySelectorAll(':scope > p'));
          if (!paragraphs.length) return;
          paragraphs.forEach(function(p, idx) {
            var frag = doc.createDocumentFragment();
            if (idx > 0) frag.appendChild(doc.createElement('br'));
            while (p.firstChild) frag.appendChild(p.firstChild);
            p.parentNode.insertBefore(frag, p);
            p.parentNode.removeChild(p);
          });
        });

        container.querySelectorAll('img').forEach(function(img) {
          var src = img.getAttribute('src') || '';
          if (src.indexOf('http://') !== 0 && src.indexOf('https://') !== 0 && src.indexOf('data:') !== 0) {
            img.parentNode.removeChild(img);
          }
        });

        container.querySelectorAll('br[style]').forEach(function(br) {
          if (/mso-/i.test(br.getAttribute('style') || '')) {
            br.removeAttribute('style');
          }
        });
      };

      this._hasOnlyNoisyStyles = function(span) {
        var style = span.getAttribute('style') || '';
        if (!style.trim()) return true;
        var VISUAL = /^(color|background-color|font-size|font-weight|font-style|text-decoration|vertical-align)\s*:/i;
        return style.split(';').map(function(p) { return p.trim(); })
          .filter(Boolean).every(function(p) { return !VISUAL.test(p); });
      };

      // -----------------------------------------------------------------------
      // Style and attribute cleaning
      // -----------------------------------------------------------------------

      this._cleanStyles = function(container) {
        var KEEP = ['color', 'background-color', 'font-size', 'font-weight',
          'font-style', 'text-decoration', 'text-align', 'vertical-align'];

        var DEFAULTS = {
          'color': ['#000000', 'black', 'windowtext', 'inherit', 'rgb(0,0,0)', 'rgb(0, 0, 0)'],
          'background-color': ['#ffffff', 'white', 'transparent', 'inherit',
            'rgb(255,255,255)', 'rgb(255, 255, 255)'],
          'font-size': ['12pt'],
          'font-weight': ['normal', '400'],
          'font-style': ['normal'],
          'vertical-align': ['baseline', 'top'],
          'text-align': ['left', 'start'],
        };

        container.querySelectorAll('[style]').forEach(function(el) {
          var cleaned = (el.getAttribute('style') || '')
            .split(';')
            .map(function(p) { return p.trim(); })
            .filter(function(p) {
              if (!p) return false;
              var colon = p.indexOf(':');
              if (colon === -1) return false;
              var prop = p.slice(0, colon).trim().toLowerCase();
              if (KEEP.indexOf(prop) === -1) return false;
              var value = p.slice(colon + 1).trim().toLowerCase()
                .replace(/\s*!important\s*$/, '');
              return !DEFAULTS[prop] || DEFAULTS[prop].indexOf(value) === -1;
            })
            .join('; ');

          if (cleaned) {
            el.setAttribute('style', cleaned);
          } else {
            el.removeAttribute('style');
          }
        });
      };

      this._cleanAttributes = function(container) {
        var PRESERVE_ON = {
          'A':   ['href', 'target', 'title', 'rel'],
          'IMG': ['src', 'alt', 'width', 'height'],
          'TD':  ['colspan', 'rowspan'],
          'TH':  ['colspan', 'rowspan', 'scope'],
          'OL':  ['start', 'type'],
        };
        var ALWAYS_KEEP = ['style'];

        container.querySelectorAll('*').forEach(function(el) {
          var tag = el.tagName.toUpperCase();
          var allowed = PRESERVE_ON[tag] || [];
          Array.from(el.attributes).forEach(function(attr) {
            var name = attr.name.toLowerCase();
            if (ALWAYS_KEEP.indexOf(name) === -1 && allowed.indexOf(name) === -1) {
              el.removeAttribute(attr.name);
            }
          });
        });
      };

      this._deduplicateInheritedStyles = function(container) {
        container.querySelectorAll('[style]').forEach(function(el) {
          var parent = el.parentElement;
          if (!parent) return;
          var parentStyles = self._parseStyleStr(parent.getAttribute('style') || '');
          if (!Object.keys(parentStyles).length) return;
          var childStyles = self._parseStyleStr(el.getAttribute('style') || '');
          var unique = Object.keys(childStyles)
            .filter(function(prop) { return parentStyles[prop] !== childStyles[prop]; })
            .map(function(prop) { return prop + ': ' + childStyles[prop]; })
            .join('; ');
          if (unique) {
            el.setAttribute('style', unique);
          } else {
            el.removeAttribute('style');
          }
        });
      };

      this._parseStyleStr = function(styleStr) {
        var result = {};
        (styleStr || '').split(';').forEach(function(decl) {
          var colon = decl.indexOf(':');
          if (colon === -1) return;
          var prop = decl.slice(0, colon).trim().toLowerCase();
          var val = decl.slice(colon + 1).trim().toLowerCase();
          if (prop && val) result[prop] = val;
        });
        return result;
      };

      this._cleanHeadingSpans = function(container) {
        Array.from(container.querySelectorAll('h1 span, h2 span, h3 span, h4 span, h5 span, h6 span'))
          .reverse()
          .forEach(function(span) {
            if (span.parentNode) span.replaceWith.apply(span, Array.from(span.childNodes));
          });
      };

      this._unwrapEmptySpans = function(container) {
        Array.from(container.querySelectorAll('span')).reverse().forEach(function(span) {
          if (span.attributes.length === 0 && span.parentNode) {
            span.replaceWith.apply(span, Array.from(span.childNodes));
          }
        });
      };

      this._unwrapWhitespaceSpans = function(container) {
        container.querySelectorAll('span').forEach(function(span) {
          if (span.textContent.trim() === '' && !span.querySelector('img, br') && span.parentNode) {
            span.replaceWith.apply(span, Array.from(span.childNodes));
          }
        });
      };

      this._replaceNbsp = function(container) {
        var walker = container.ownerDocument.createTreeWalker(container, 4 /* NodeFilter.SHOW_TEXT */);
        var node;
        while ((node = walker.nextNode())) {
          if (node.nodeValue.indexOf('\u00a0') !== -1) {
            node.nodeValue = node.nodeValue.replace(/\u00a0/g, ' ');
          }
        }
      };

      this._removeEmptyBlocks = function(container) {
        var changed = true;
        while (changed) {
          changed = false;
          container.querySelectorAll('p, div, h1, h2, h3, h4, h5, h6, span').forEach(function(el) {
            if (!el.textContent.trim() && !el.querySelector('img, br')) {
              el.parentNode.removeChild(el);
              changed = true;
            }
          });
        }
      };
    },
  });
}));
