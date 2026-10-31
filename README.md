# MCQ Slide Maker

Browser-e chole, free GitHub Pages e host hoy. Template (.pptx) upload kor, MCQ paste kor, prottek question-er alada slide shoho PPTX download kor. Kono server ba API key lage na.

## Kivabe kaj kore
1. Template-e ekta **question sample slide** thakbe ("(A) … (B) …" ar "Ans:" shoho) ar chaile ekta **"Type…" heading sample slide**.
2. Site oi sample slide duplicate kore prottek question-er text boshay, box-er height adjust kore, tarpor sample duto remove kore.
3. Cover ba onno slide gulo jemon chhilo temon thake.

## Input format
```
T|Type-01: Heading
Q|Question text [DU. 08-09]|Option A|Option B|Option C|Option D|B
```
Optional 8th field: extra line (option-er niche).

## Local e chalano
```
python3 -m http.server 8000   # tarpor http://localhost:8000
```
## Test (Node)
```
npm i jszip @xmldom/xmldom
node test/run.js templates/demo-rtd.pptx samples/sample.txt out.pptx
```
## GitHub Pages
Repo Settings → Pages → Branch `main` / root.

> Copyrighted boi/PDF-er question repo-te commit korish na, shudhu nijer ba free-to-use content rakh.
