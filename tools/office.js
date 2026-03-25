const OFFICE_BRIDGE_URL = process.env.OPENCODE_OFFICE_BRIDGE_URL || 'http://127.0.0.1:52391/api/office-tools/execute';

async function execute(host, toolName, args) {
  const response = await fetch(OFFICE_BRIDGE_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ host, toolName, args }),
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
  headingText: { type: 'string', description: 'Heading text to search for.' },
  includeSubsections: { type: 'boolean', description: 'Include nested subsections.' },
});
exports.set_document_content = word('set_document_content', 'Replace the current Word document with new HTML content.', {
  html: { type: 'string', description: 'HTML to write into the document.' },
});
exports.get_selection = word('get_selection', 'Read the current Word selection as OOXML.', {});
exports.get_selection_text = word('get_selection_text', 'Read the current Word selection as plain text.', {});
exports.insert_content_at_selection = word('insert_content_at_selection', 'Insert HTML content at the current Word selection.', {
  html: { type: 'string', description: 'HTML to insert.' },
  location: { type: 'string', description: 'Where to insert relative to the current selection.' },
});
exports.find_and_replace = word('find_and_replace', 'Find and replace text throughout the active Word document.', {
  find: { type: 'string', description: 'Text to find.' },
  replace: { type: 'string', description: 'Replacement text.' },
  matchCase: { type: 'boolean', description: 'Match case exactly.' },
  matchWholeWord: { type: 'boolean', description: 'Only match whole words.' },
});
exports.insert_table = word('insert_table', 'Insert a table at the current Word selection.', {
  data: { type: 'array', description: 'Two-dimensional array of table rows.' },
  hasHeader: { type: 'boolean', description: 'Treat the first row as a header row.' },
  style: { type: 'string', description: 'Table style.' },
});
exports.apply_style_to_selection = word('apply_style_to_selection', 'Apply formatting styles to the current Word selection.', {
  bold: { type: 'boolean' },
  italic: { type: 'boolean' },
  underline: { type: 'boolean' },
  strikethrough: { type: 'boolean' },
  fontSize: { type: 'number' },
  fontName: { type: 'string' },
  fontColor: { type: 'string' },
  highlightColor: { type: 'string' },
});
