// node test/run.js <template.pptx> <data.txt> <out.pptx>
const fs = require('fs'); const MCQ = require('../core.js');
(async () => {
  const [tpl, data, out] = process.argv.slice(2);
  const { items, errors } = MCQ.parseItems(fs.readFileSync(data, 'utf8'));
  if (errors.length) console.log(errors);
  console.log(await MCQ.inspect(fs.readFileSync(tpl)).then(i => ({ ...i, texts: undefined })));
  const bytes = await MCQ.generate(fs.readFileSync(tpl), items);
  fs.writeFileSync(out, bytes); console.log('items', items.length, 'bytes', bytes.length);
})();
