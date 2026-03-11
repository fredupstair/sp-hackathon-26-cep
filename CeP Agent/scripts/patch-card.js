/**
 * patch-card.js
 * Injects response_semantics (adaptive card) into the TypeSpec-generated
 * cepapi-apiplugin.json after every typeSpec/compile run.
 *
 * Called via "npm run patch:card" in m365agents.yml / m365agents.local.yml.
 */

const fs = require('fs');
const path = require('path');

const generatedDir = path.resolve(__dirname, '../appPackage/.generated');
const pluginPath   = path.join(generatedDir, 'cepapi-apiplugin.json');
const cardPath     = path.join(generatedDir, 'adaptiveCards/postWin.json');

const plugin = JSON.parse(fs.readFileSync(pluginPath, 'utf8'));
const card   = JSON.parse(fs.readFileSync(cardPath, 'utf8'));

const postWin = plugin.functions && plugin.functions.find(f => f.name === 'postWin');
if (!postWin) {
  console.error('patch-card.js: postWin function not found in', pluginPath);
  process.exit(1);
}

postWin.capabilities = {
  response_semantics: {
    data_path: '$',
    static_template: card
  }
};

fs.writeFileSync(pluginPath, JSON.stringify(plugin, null, 4), 'utf8');
console.log('✓ patch-card.js: injected adaptive card into', pluginPath);
