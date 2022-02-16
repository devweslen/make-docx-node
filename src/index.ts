import * as fs from 'fs';
import {
  File,
  HeadingLevel,
  ISectionOptions,
  Packer, Paragraph,
  StyleLevel, Table,
  TableCell, TableOfContents,
  TableRow, WidthType,
} from 'docx';

const page1 = {
  children: [
    new TableOfContents('Summary', {
      hyperlink: true,
      headingStyleRange: '1-5',
      stylesWithLevels: [new StyleLevel('MySpectacularStyle', 1)],
    }),
  ],
} as ISectionOptions;

const page2 = {
  children: [
    new Paragraph({
      text: 'Introdução',
      heading: HeadingLevel.HEADING_1,
      numbering: {
        level: 0,
        reference: '%2.',
      },
    }),
    new Paragraph("I'm a other text very nicely written.'"),
  ],
} as ISectionOptions;

const table = new Table({
  columnWidths: [3505, 5505],
  rows: [
    new TableRow({
      children: [
        new TableCell({
          width: {
            size: 3505,
            type: WidthType.DXA,
          },
          children: [new Paragraph('Hello')],
        }),
        new TableCell({
          width: {
            size: 5505,
            type: WidthType.DXA,
          },
          children: [],
        }),
      ],
    }),
    new TableRow({
      children: [
        new TableCell({
          width: {
            size: 3505,
            type: WidthType.DXA,
          },
          children: [],
        }),
        new TableCell({
          width: {
            size: 5505,
            type: WidthType.DXA,
          },
          children: [new Paragraph('World')],
        }),
      ],
    }),
  ],
});

const page3 = {
  children: [
    new Paragraph({
      text: 'Desenvolvimento',
      heading: HeadingLevel.HEADING_1,
    }),
    table,
  ],
} as ISectionOptions;

const doc = new File({
  features: {
    updateFields: true,
  },
  sections: [
    page1,
    page2,
    page3,
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  const now = new Date().getTime();
  fs.writeFileSync(`./src/docs/${now}.docx`, buffer);
});

console.log('DOCUMENTO CRIADO');
