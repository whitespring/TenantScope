import assert from 'node:assert/strict';
import { createDocx } from './docx.mjs';

const bytes = createDocx('Test & Bericht', [
  { type: 'meta', rows: [['Organisation', 'Contoso'], ['Erstellt', '3.9.2026']] },
  { type: 'metrics', items: [{ value: '72/100', label: 'Governance-Indikator' }, { value: '2', label: 'Hohes Risiko' }] },
  { type: 'finding', severity: 'high', label: 'Hoch', title: 'Testbefund', description: 'Feststellung', context: 'Einordnung', goodPractice: 'Good Practice', action: 'Jetzt prüfen', links: [{ text: 'Microsoft-Hilfe', href: 'https://learn.microsoft.com/' }] },
  { type: 'table', columns: ['Bereich', 'Status'], rows: [['Identitäten', 'Live analysiert']], widths: [6000, 3360], alignments: ['left', 'center'] },
  { type: 'callout', title: 'Identitäten', text: 'Regelmäßig rezertifizieren.', links: [{ text: 'Admin Center', href: 'https://entra.microsoft.com/' }] },
  { text: 'Tenant <Nord>', style: 'Heading1' },
  { text: 'Microsoft-Hilfe', href: 'https://learn.microsoft.com/?a=1&b=2' },
  { text: 'Nicht verlinken', href: 'https://example.invalid/tracker' },
]);

assert.equal(new DataView(bytes.buffer).getUint32(0, true), 0x04034b50, 'DOCX must start with a ZIP header');
assert.ok(new TextDecoder().decode(bytes).includes('Tenant &lt;Nord&gt;'), 'document text must be XML-escaped');
assert.ok(new TextDecoder().decode(bytes).includes('relationships/hyperlink'), 'DOCX links must be real hyperlinks');
assert.ok(new TextDecoder().decode(bytes).includes('a=1&amp;b=2'), 'DOCX hyperlink targets must be XML-escaped');
assert.ok(!new TextDecoder().decode(bytes).includes('example.invalid'), 'DOCX must not create links to untrusted hosts');
assert.ok(new TextDecoder().decode(bytes).includes('<w:tblGrid>'), 'DOCX boxes and tables must use explicit Word table geometry');
assert.ok(new TextDecoder().decode(bytes).includes('<w:tblHeader/>'), 'DOCX data-table headers must repeat across pages');
assert.ok(new TextDecoder().decode(bytes).includes('w:fill="FDE8E7"'), 'DOCX findings must retain severity colors');
assert.ok(new TextDecoder().decode(bytes).includes('header1.xml'), 'DOCX must include report header and footer furniture');
assert.ok(new TextDecoder().decode(bytes).includes('xml:space="preserve">Seite '), 'DOCX page-number label must retain its trailing space');
assert.ok(bytes.length > 1000, 'DOCX package must contain its required XML parts');

console.log('DOCX export self-check passed');
