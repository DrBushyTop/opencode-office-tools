const { z } = require('zod');

const OFFICE_BRIDGE_URL = process.env.OPENCODE_OFFICE_BRIDGE_URL || 'http://127.0.0.1:52391/api/office-tools/execute';

async function execute(host, tool, args) {
  const response = await fetch(OFFICE_BRIDGE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ host, toolName: tool, args }),
  });

  if (!response.ok) {
    throw new Error((await response.text()) || `Office tool failed: ${response.status}`);
  }

  const result = await response.json();
  if (typeof result.result === 'string') return result.result;
  return JSON.stringify(result.result, null, 2);
}

function word(name, description, args) {
  return {
    description,
    args,
    async execute(input) {
      return execute('word', name, input);
    },
  };
}

exports.get_document_overview = word('get_document_overview', 'Get a structural overview of the active Word document.', {});
exports.get_document_content = word('get_document_content', 'Read the current Word document.', {});
exports.get_document_section = word('get_document_section', 'Read a specific Word document section by heading.', {
  headingText: z.string().describe('Heading text to search for.'),
  includeSubsections: z.boolean().optional().describe('Include nested subsections.'),
});
exports.set_document_content = word('set_document_content', 'Replace the current Word document with new HTML content.', {
  html: z.string().describe('HTML to write into the document.'),
});
exports.get_selection = word('get_selection', 'Read the current Word selection as OOXML.', {});
exports.get_selection_text = word('get_selection_text', 'Read the current Word selection as plain text.', {});
exports.insert_content_at_selection = word('insert_content_at_selection', 'Insert HTML content at the current Word selection.', {
  html: z.string().describe('HTML to insert.'),
  location: z.enum(['replace', 'before', 'after', 'start', 'end']).optional().describe('Where to insert relative to the current selection.'),
});
exports.find_and_replace = word('find_and_replace', 'Find and replace text throughout the active Word document.', {
  find: z.string().describe('Text to find.'),
  replace: z.string().describe('Replacement text.'),
  matchCase: z.boolean().optional().describe('Match case exactly.'),
  matchWholeWord: z.boolean().optional().describe('Only match whole words.'),
});
exports.insert_table = word('insert_table', 'Insert a table at the current Word selection.', {
  data: z.array(z.array(z.string())).describe('Two-dimensional array of table cell values.'),
  hasHeader: z.boolean().optional().describe('Treat the first row as a header row.'),
  style: z.enum(['grid', 'striped', 'plain']).optional().describe('Table style.'),
});
exports.apply_style_to_selection = word('apply_style_to_selection', 'Apply formatting styles to the current Word selection.', {
  bold: z.boolean().optional(),
  italic: z.boolean().optional(),
  underline: z.boolean().optional(),
  strikethrough: z.boolean().optional(),
  fontSize: z.number().optional(),
  fontName: z.string().optional(),
  fontColor: z.string().optional(),
  highlightColor: z.string().optional(),
});
