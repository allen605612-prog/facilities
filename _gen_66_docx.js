const {
  Document, Packer, Paragraph, TextRun, AlignmentType
} = require('docx');
const fs = require('fs');

const font = '標楷體';
const sz = 28; // 14pt = 28 half-points

function makeParagraph(text) {
  return new Paragraph({
    children: [
      new TextRun({
        text,
        font: { name: font, eastAsia: font },
        bold: true,
        size: sz,
      })
    ]
  });
}

function emptyParagraph() {
  return new Paragraph({
    children: [
      new TextRun({
        text: '',
        font: { name: font, eastAsia: font },
        bold: true,
        size: sz,
      })
    ]
  });
}

// 66屆得獎名單
const entries = [
  {
    students: '本校美二甲班廖采玲、陳彙喬、黃柏之',
    award: '參加雲林縣第66屆科展榮獲國中數學科第一名',
    teacher: '指導老師：羅婉萍',
  },
  {
    students: '美一乙班廖泱泱',
    award: '榮獲國中數學科第二名',
    teacher: '指導老師：詩富強',
  },
  {
    students: '音二甲班洪靖宇、蔡承恩',
    award: '榮獲國中生物科第二名',
    teacher: '指導老師：許議鶴',
  },
  {
    students: '美二乙班黃家恩、美二甲班王嘉裕、音二甲班陳昱勳',
    award: '榮獲國中生物科第三名',
    teacher: '指導老師：郭靜如',
  },
  {
    students: '初一誠班陳沛青、美一甲班鍾子鵬、美一乙班陳意凡',
    award: '榮獲國中生活與應用科學科(一)第四名',
    teacher: '指導老師：林俊杰',
  },
  {
    students: '音二甲班籃苡熏、萬奕岑、陳姸汝',
    award: '榮獲國中生活與應用科學科(二)第三名',
    teacher: '指導老師：許議鶴',
  },
  {
    students: '初二正班許呈睿、黃品語、洪啟鈞',
    award: '榮獲國中生活與應用科學科(三)第四名',
    teacher: '指導老師：許議鶴',
  },
];

const children = [emptyParagraph()];

for (const entry of entries) {
  children.push(makeParagraph(entry.students));
  children.push(makeParagraph(`${entry.award}  ${entry.teacher}`));
  children.push(emptyParagraph());
}

children.push(makeParagraph('恭喜以上獲獎同學。'));

const doc = new Document({
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1440, right: 1800, bottom: 1440, left: 1800 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('D:\\D\\113設備組\\科展頒獎\\66屆縣科展頒獎名單.docx', buffer);
  console.log('完成：66屆縣科展頒獎名單.docx');
});
