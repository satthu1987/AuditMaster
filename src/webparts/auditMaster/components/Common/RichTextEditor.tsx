import * as React from 'react';
import ReactQuill from 'react-quill';
// @ts-ignore: side-effect CSS import has no type declarations
import 'react-quill/dist/quill.snow.css';
import { Label, mergeStyles } from '@fluentui/react';
import styles from './RichTextEditor.module.scss';

// Quill is exposed as a static member of the ReactQuill class (no named export).
const Quill: any = ReactQuill.Quill;

// ════════════════════════════════════════════════════════════════════════════
//  Table support for Quill 1.3.7
//  Quill 1.3.7 ships WITHOUT a built-in table module, so we register a small set
//  of container blots (table → tbody → tr → td) once. This lets pasted/inserted
//  table HTML survive Quill's content model AND keeps the individual cells fully
//  editable inline. Registration is guarded so it only happens a single time.
// ════════════════════════════════════════════════════════════════════════════

let blotsRegistered = false;

function registerTableBlots(): void {
  if (blotsRegistered) return;

  try {
    const Block: any = Quill.import('blots/block');
    const Container: any = Quill.import('blots/container');

    // <td> – a block-level editable cell
    class TableCell extends Block {
      public static blotName = 'table-cell';
      public static tagName = 'TD';
    }

    // <tr> – row containing cells
    class TableRow extends Container {
      public static blotName = 'table-row';
      public static tagName = 'TR';
    }
    (TableRow as any).allowedChildren = [TableCell];

    // <tbody> – body containing rows
    class TableBody extends Container {
      public static blotName = 'table-body';
      public static tagName = 'TBODY';
    }
    (TableBody as any).allowedChildren = [TableRow];

    // <table> – top-level container
    class TableContainer extends Container {
      public static blotName = 'table';
      public static tagName = 'TABLE';
    }
    (TableContainer as any).allowedChildren = [TableBody];

    Quill.register(TableCell, true);
    Quill.register(TableRow, true);
    Quill.register(TableBody, true);
    Quill.register(TableContainer, true);

    blotsRegistered = true;
  } catch (e) {
    // If registration fails (e.g. blots already exist), don't block the editor.
    // eslint-disable-next-line no-console
    console.warn('RichTextEditor: table blot registration skipped:', e);
    blotsRegistered = true;
  }
}

// ─── Helpers ─────────────────────────────────────────────────────────────────

/**
 * Quill emits placeholder markup such as `<p><br></p>` for an "empty" editor.
 * We normalise that to an empty string so SharePoint stores clean data and
 * required-field validation behaves correctly.
 */
function normalizeHtml(html: string): string {
  if (!html) return '';
  const stripped = html
    .replace(/<p>(\s|&nbsp;|<br\s*\/?>)*<\/p>/gi, '')
    .replace(/<br\s*\/?>/gi, '')
    .replace(/&nbsp;/gi, '')
    .replace(/\s/g, '');
  return stripped.length === 0 ? '' : html;
}

// ─── Public interface ────────────────────────────────────────────────────────

export interface IRichTextEditorProps {
  /** Field label shown above the editor */
  label?: string;
  /** Current HTML content */
  value: string;
  /** Callback when content changes – returns sanitized HTML string */
  onChange: (html: string) => void;
  /** When true, the editor is read-only */
  disabled?: boolean;
  /** Whether the field is required */
  required?: boolean;
  /** Placeholder text shown when empty */
  placeholder?: string;
  /** Minimum editor height in px (default 160) */
  minHeight?: number;
  /** Optional description text rendered below the editor */
  description?: string;
}

// ─── Constants ───────────────────────────────────────────────────────────────

// A unique toolbar id is required so multiple editors on the same page do not
// share/clash toolbars.
let editorCounter = 0;

const wrapperClass = mergeStyles({
  selectors: {
    '.ql-toolbar.ql-snow': {
      borderTopLeftRadius: 2,
      borderTopRightRadius: 2,
      borderColor: '#8a8886',
      fontFamily: '"Segoe UI", "Segoe UI Web (West European)", sans-serif'
    },
    '.ql-container.ql-snow': {
      borderColor: '#8a8886',
      borderBottomLeftRadius: 2,
      borderBottomRightRadius: 2,
      fontFamily: '"Segoe UI", "Segoe UI Web (West European)", sans-serif',
      fontSize: 14
    },
    '.ql-editor table': {
      borderCollapse: 'collapse',
      width: '100%',
      margin: '8px 0'
    },
    '.ql-editor td': {
      border: '1px solid #c8c6c4',
      padding: '6px 8px',
      minWidth: 40
    }
  }
});

// ─── Component ───────────────────────────────────────────────────────────────

/**
 * Reusable rich-text editor built on react-quill (Quill 1.3.7).
 *
 * Features:
 *  - Headers, bold, italic, underline, strike
 *  - Ordered / bullet lists, indent
 *  - Text & background color pickers
 *  - Text alignment
 *  - Link insertion (Quill default prompt)
 *  - Image insertion (custom URL prompt – keeps content lightweight)
 *  - Table insertion (custom rows × cols dialog, editable cells)
 *  - Clean / remove-formatting button
 *  - Read-only mode (toolbar hidden, content not editable)
 *
 * HTML content round-trips cleanly to SharePoint "Multiple lines of text"
 * (enhanced rich text) columns.
 */
export const RichTextEditor: React.FC<IRichTextEditorProps> = (props) => {
  const {
    label,
    value,
    onChange,
    disabled = false,
    required = false,
    placeholder = 'Enter description...',
    minHeight = 160,
    description
  } = props;

  const quillRef = React.useRef<ReactQuill | null>(null);

  // Stable, unique toolbar id for this editor instance
  const toolbarId = React.useMemo(() => {
    editorCounter += 1;
    return `rte-toolbar-${editorCounter}`;
  }, []);

  // Register the table blots before the first editor mounts
  React.useMemo(() => registerTableBlots(), []);

  // ── Custom image handler – prompt for a URL ──────────────────────────────
  const imageHandler = React.useCallback(() => {
    const editor = quillRef.current?.getEditor();
    if (!editor) return;
    const url = window.prompt('Enter the image URL:');
    if (url) {
      const range = editor.getSelection(true);
      editor.insertEmbed(range ? range.index : 0, 'image', url, 'user');
      editor.setSelection((range ? range.index : 0) + 1, 0);
    }
  }, []);

  // ── Custom table handler – prompt for rows × cols ────────────────────────
  const tableHandler = React.useCallback(() => {
    const editor = quillRef.current?.getEditor();
    if (!editor) return;

    const rowsRaw = window.prompt('Number of rows:', '2');
    if (!rowsRaw) return;
    const colsRaw = window.prompt('Number of columns:', '2');
    if (!colsRaw) return;

    const rows = Math.max(1, Math.min(20, parseInt(rowsRaw, 10) || 0));
    const cols = Math.max(1, Math.min(10, parseInt(colsRaw, 10) || 0));
    if (!rows || !cols) return;

    let html = '<table><tbody>';
    for (let r = 0; r < rows; r++) {
      html += '<tr>';
      for (let c = 0; c < cols; c++) {
        html += '<td><br></td>';
      }
      html += '</tr>';
    }
    html += '</tbody></table><p><br></p>';

    const range = editor.getSelection(true);
    const index = range ? range.index : editor.getLength();
    editor.clipboard.dangerouslyPasteHTML(index, html, 'user');
  }, []);

  // ── Modules config (memoised) ────────────────────────────────────────────
  const modules = React.useMemo(
    () => ({
      toolbar: {
        container: `#${toolbarId}`,
        handlers: {
          image: imageHandler,
          'insert-table': tableHandler
        }
      },
      clipboard: {
        matchVisual: false
      }
    }),
    [toolbarId, imageHandler, tableHandler]
  );

  const formats = React.useMemo(
    () => [
      'header',
      'bold',
      'italic',
      'underline',
      'strike',
      'color',
      'background',
      'list',
      'bullet',
      'indent',
      'align',
      'link',
      'image',
      'table',
      'table-body',
      'table-row',
      'table-cell'
    ],
    []
  );

  // ── Read-only render ──────────────────────────────────────────────────────
  if (disabled) {
    return (
      <div className={styles.rteContainer}>
        {label && <Label required={required}>{label}</Label>}
        <div
          className={styles.readOnlyContent}
          style={{ minHeight }}
          // Content is HTML produced by Quill; rendered read-only.
          dangerouslySetInnerHTML={{ __html: value || '<span class="rte-empty">—</span>' }}
        />
        {description && <span className={styles.description}>{description}</span>}
      </div>
    );
  }

  // ── Editable render ───────────────────────────────────────────────────────
  return (
    <div className={`${styles.rteContainer} ${wrapperClass}`}>
      {label && <Label required={required}>{label}</Label>}

      {/* Custom toolbar – referenced by modules.toolbar.container */}
      <div id={toolbarId}>
        <span className="ql-formats">
          <select className="ql-header" defaultValue="">
            <option value="1">Heading 1</option>
            <option value="2">Heading 2</option>
            <option value="3">Heading 3</option>
            <option value="">Normal</option>
          </select>
        </span>
        <span className="ql-formats">
          <button className="ql-bold" title="Bold" />
          <button className="ql-italic" title="Italic" />
          <button className="ql-underline" title="Underline" />
          <button className="ql-strike" title="Strikethrough" />
        </span>
        <span className="ql-formats">
          <select className="ql-color" title="Text color" />
          <select className="ql-background" title="Highlight color" />
        </span>
        <span className="ql-formats">
          <button className="ql-list" value="ordered" title="Numbered list" />
          <button className="ql-list" value="bullet" title="Bulleted list" />
          <button className="ql-indent" value="-1" title="Decrease indent" />
          <button className="ql-indent" value="+1" title="Increase indent" />
        </span>
        <span className="ql-formats">
          <select className="ql-align" title="Alignment" />
        </span>
        <span className="ql-formats">
          <button className="ql-link" title="Insert link" />
          <button className="ql-image" title="Insert image (URL)" />
          <button className="ql-insert-table" title="Insert table">
            {/* Simple table glyph */}
            <svg viewBox="0 0 18 18" width="18" height="18">
              <rect className="ql-stroke" height="12" width="12" x="3" y="3" fill="none" />
              <line className="ql-stroke" x1="3" y1="7" x2="15" y2="7" />
              <line className="ql-stroke" x1="3" y1="11" x2="15" y2="11" />
              <line className="ql-stroke" x1="9" y1="3" x2="9" y2="15" />
            </svg>
          </button>
        </span>
        <span className="ql-formats">
          <button className="ql-clean" title="Remove formatting" />
        </span>
      </div>

      <ReactQuill
        ref={quillRef}
        theme="snow"
        value={value || ''}
        onChange={(content) => onChange(normalizeHtml(content))}
        modules={modules}
        formats={formats}
        placeholder={placeholder}
        readOnly={disabled}
        style={{ minHeight }}
      />

      {description && <span className={styles.description}>{description}</span>}
    </div>
  );
};

export default RichTextEditor;
