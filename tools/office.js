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

exports.get_document_content = {
  description: 'Read the current Word document.',
  args: {},
  async execute(args) {
    return execute('word', 'get_document_content', args);
  },
};
