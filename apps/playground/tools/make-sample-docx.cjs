// Generates apps/playground/public/sample.docx — a multi-page document that
// exercises the import + layout + paint pipeline end to end: marks (bold,
// italic, color, size, underline, strike, highlight, super/subscript),
// alignment (center/right/justify), first-line and hanging indent, paragraph
// spacing (w:spacing) + soft line breaks (w:br) + symbol runs (w:sym) +
// page-break-before, multi-level numbered + bullet lists, tables (colwidth,
// colspan, rowspan/vMerge, table align, row height, cell vAlign, per-cell &
// table borders, shading), inline + floating images, hyperlinks, tab stops.
//
//   node apps/playground/tools/make-sample-docx.cjs
const fs = require('node:fs');
const path = require('node:path');
const JSZip = require('jszip');

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const WP_NS = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PIC_NS = 'http://schemas.openxmlformats.org/drawingml/2006/picture';

// 1x1 PNGs; drawn at their wp:extent size (px × 9525 EMU).
const PNG_RED =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const PNG_BLUE =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPj/HwADBwIAMCbHYQAAAABJRU5ErkJggg==';

const LOREM = [
  'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat.',
  'Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.',
  'Sed ut perspiciatis unde omnis iste natus error sit voluptatem accusantium doloremque laudantium, totam rem aperiam, eaque ipsa quae ab illo inventore veritatis et quasi architecto beatae vitae dicta sunt explicabo.',
  'Nemo enim ipsam voluptatem quia voluptas sit aspernatur aut odit aut fugit, sed quia consequuntur magni dolores eos qui ratione voluptatem sequi nesciunt, neque porro quisquam est qui dolorem ipsum.',
  'Việt Nam đất nước ta ơi, mênh mông biển lúa đâu trời đẹp hơn. Cánh cò bay lả rập rờn, mây mờ che đỉnh Trường Sơn sớm chiều. Quê hương biết mấy thân yêu, bao nhiêu đời đã chịu nhiều thương đau.',
  'Gõ tiếng Việt có dấu là bài kiểm tra quan trọng nhất của bapbong: chữ ư, ơ, ă, â, ê, ô cùng các thanh sắc huyền hỏi ngã nặng phải hiển thị đúng trên canvas ở mọi cỡ chữ và kiểu chữ.',
];

const esc = (s) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const run = (text, rpr = '') =>
  `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${esc(text)}</w:t></w:r>`;
const p = (content, ppr = '') => `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ''}${content}</w:p>`;
const jc = (v) => `<w:jc w:val="${v}"/>`;
const listP = (numId, ilvl, text) =>
  p(run(text), `<w:numPr><w:ilvl w:val="${ilvl}"/><w:numId w:val="${numId}"/></w:numPr>`);
const td = (content, tcPr = '') => `<w:tc><w:tcPr>${tcPr}</w:tcPr>${content}</w:tc>`;
const image = (relId, sizePx, alt) =>
  `<w:r><w:drawing><wp:inline><wp:extent cx="${sizePx * 9525}" cy="${sizePx * 9525}"/><wp:docPr id="1" name="${alt}" descr="${alt}"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="${relId}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

const heading = (text) => p(run(text, '<w:b/><w:sz w:val="28"/>'), jc('left'));

// Lưới viền đầy đủ — OOXML mặc định KHÔNG có viền; bảng nào muốn lưới phải khai.
const TBL_BORDERS = `<w:tblPr><w:tblBorders><w:top w:val="single"/><w:bottom w:val="single"/><w:left w:val="single"/><w:right w:val="single"/><w:insideH w:val="single"/><w:insideV w:val="single"/></w:tblBorders></w:tblPr>`;

/** A chapter = heading + body paragraphs + one "special" showcase element. */
function chapter(n, special) {
  return [
    heading(`${n}. ${special.title}`),
    p(run(LOREM[(n * 2) % LOREM.length]), jc('both')),
    special.xml,
    p(run(LOREM[(n * 2 + 1) % LOREM.length]), jc('both')),
  ].join('\n');
}

const SPECIALS = [
  {
    title: 'Chữ đậm, nghiêng, màu và cỡ chữ',
    xml: p(
      [
        run('Đoạn này trộn '),
        run('đậm', '<w:b/>'),
        run(', '),
        run('nghiêng', '<w:i/>'),
        run(', '),
        run('đậm nghiêng', '<w:b/><w:i/>'),
        run(', '),
        run('màu đỏ', '<w:color w:val="C0392B"/>'),
        run(', '),
        run('màu xanh', '<w:color w:val="1F6FEB"/>'),
        run(', '),
        run('gạch chân', '<w:u w:val="single"/>'),
        run(', '),
        run('gạch ngang', '<w:strike/>'),
        run(', '),
        run('tô vàng', '<w:highlight w:val="yellow"/>'),
        run(', công thức E=mc'),
        run('2', '<w:vertAlign w:val="superscript"/>'),
        run(' và H'),
        run('2', '<w:vertAlign w:val="subscript"/>'),
        run('O (super/subscript), '),
        run('chữ to 14pt', '<w:sz w:val="28"/>'),
        run(' và '),
        run('chữ nhỏ 8pt', '<w:sz w:val="16"/>'),
        run(' trong cùng một dòng để kiểm tra đo chữ theo từng run.'),
      ].join(''),
      `<w:ind w:firstLine="720"/>`,
    ),
  },
  {
    title: 'Danh sách đa cấp',
    xml: [
      listP('1', 0, 'Hạng mục thứ nhất của danh sách đánh số'),
      listP('1', 1, 'Mục con a — thụt lề cấp hai'),
      listP('1', 1, 'Mục con b đủ dài để tự xuống dòng và kiểm tra hanging indent của marker khi nội dung tràn sang dòng tiếp theo'),
      listP('1', 0, 'Hạng mục thứ hai quay về cấp một'),
      listP('2', 0, 'Gạch đầu dòng thứ nhất'),
      listP('2', 0, 'Gạch đầu dòng thứ hai'),
    ].join('\n'),
  },
  {
    title: 'Bảng: căn giữa, gộp ô, nền & viền ô',
    // Bảng căn giữa (w:jc), hàng tiêu đề cao 0.4in + canh dọc giữa (w:vAlign),
    // nền xanh; một ô body khai viền riêng (w:tcBorders).
    xml: `<w:tbl>
      <w:tblPr>
        <w:jc w:val="center"/>
        <w:tblBorders><w:top w:val="single"/><w:bottom w:val="single"/><w:left w:val="single"/><w:right w:val="single"/><w:insideH w:val="single"/><w:insideV w:val="single"/></w:tblBorders>
      </w:tblPr>
      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2800"/><w:gridCol w:w="2800"/></w:tblGrid>
      <w:tr>
        <w:trPr><w:trHeight w:val="600" w:hRule="atLeast"/></w:trPr>
        ${td(p(run('Khu vực', '<w:b/>')), '<w:shd w:fill="D9E2F3"/><w:vAlign w:val="center"/>')}
        ${td(p(run('Quý 1', '<w:b/>')), '<w:shd w:fill="D9E2F3"/><w:vAlign w:val="center"/>')}
        ${td(p(run('Quý 2', '<w:b/>')), '<w:shd w:fill="D9E2F3"/><w:vAlign w:val="center"/>')}
      </w:tr>
      <w:tr>
        ${td(p(run('Miền Bắc — ô gộp dọc qua hai hàng')), '<w:vMerge w:val="restart"/>')}
        ${td(p(run('120 đơn hàng')))}
        ${td(p(run('158 đơn hàng, tăng trưởng tốt nhờ chiến dịch mới')))}
      </w:tr>
      <w:tr>
        ${td('<w:p/>', '<w:vMerge/>')}
        ${td(p(run('95')), '<w:tcBorders><w:top w:val="nil"/><w:bottom w:val="nil"/></w:tcBorders>')}
        ${td(p(run('102')))}
      </w:tr>
      <w:tr>
        ${td(p(run('Tổng hợp cả nước — ô gộp ngang ba cột, canh dưới', '<w:i/>')), '<w:gridSpan w:val="3"/><w:vAlign w:val="bottom"/>')}
      </w:tr>
    </w:tbl>`,
  },
  {
    title: 'Ảnh inline và canh lề',
    xml: [
      p(run('Ảnh nhỏ ') + image('rId7', 32, 'red square') + run(' nằm giữa chữ, còn ảnh to hơn ') + image('rId8', 72, 'blue square') + run(' đẩy chiều cao dòng lên theo.')),
      p(run('Đoạn này canh giữa.'), jc('center')),
      p(run('Đoạn này canh phải.'), jc('right')),
    ].join('\n'),
  },
  {
    title: 'Liên kết và thụt lề treo',
    xml: [
      p(`<w:hyperlink r:id="rId9">${run('Trang chủ ProseMirror')}</w:hyperlink>${run(' — hyperlink qua relationship, mark link giữ href trong model.')}`),
      p(
        run('Thụt lề treo: dòng đầu lùi ra ngoài, các dòng tiếp theo thẳng hàng với lề trái đã thụt — kiểu trình bày danh mục tài liệu tham khảo quen thuộc trong văn bản học thuật.'),
        `<w:ind w:left="720" w:hanging="720"/>`,
      ),
      p(
        run('Đoạn này dùng style "Quote": thụt lề hai bên và chữ nghiêng đến từ STYLE (cascade pPr qua w:basedOn), không khai inline — đúng cách văn bản công sở dùng Heading/Quote/ListParagraph.'),
        `<w:pStyle w:val="Quote"/>`,
      ),
    ].join('\n'),
  },
  {
    title: 'Ảnh floating — chữ chảy quanh ảnh',
    xml: [
      p(
        `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300">
          <wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>
          <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
          <wp:extent cx="1333500" cy="1333500"/>
          <wp:docPr id="2" name="Float 1" descr="blue float"/>
          <wp:wrapSquare wrapText="bothSides"/>
          <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId8"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>
        </wp:anchor></w:drawing></w:r>` +
          run('Ảnh xanh 140px neo bên phải với wrapSquare: dòng nào nằm ngang ảnh sẽ thu hẹp lại, chữ chảy dọc theo mép trái của ảnh. ') +
          run(LOREM[0]) +
          run(' ' + LOREM[4]),
      ),
      p(run('Đoạn kế tiếp không có ảnh nhưng vẫn né rect của ảnh nếu ảnh còn thò xuống — exclusion sống theo trang, không theo đoạn. ' + LOREM[5])),
      p(
        `<w:r><w:drawing><wp:anchor distT="57150" distB="57150">
          <wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>
          <wp:positionV relativeFrom="paragraph"><wp:posOffset>190500</wp:posOffset></wp:positionV>
          <wp:extent cx="952500" cy="476250"/>
          <wp:docPr id="3" name="Float 2" descr="red band"/>
          <wp:wrapTopAndBottom/>
          <a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId7"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>
        </wp:anchor></w:drawing></w:r>` +
          run('Còn ảnh đỏ dưới đây dùng wrapTopAndBottom: không dòng nào đứng cạnh ảnh — văn bản nhảy hẳn xuống dưới. ' + LOREM[1]),
      ),
    ].join('\n'),
  },
  {
    title: 'Tab stops tùy biến (w:tabs)',
    xml: [
      // Mục lục: right-tab + dot leader sát lề phải.
      ...[
        ['Chương 1 — Mở đầu', '3'],
        ['Chương 2 — Kiến trúc canvas', '12'],
        ['Chương 3 — Ngắt trang bảng', '27'],
      ].map(([title, page]) =>
        p(
          `${run(title)}<w:r><w:tab/></w:r>${run(page)}`,
          `<w:tabs><w:tab w:val="right" w:pos="9000" w:leader="dot"/></w:tabs>`,
        ),
      ),
      // Chữ ký hai bên: center-tab giữa trang + right-tab sát lề.
      p(
        `${run('NGƯỜI LẬP', '<w:b/>')}<w:r><w:tab/></w:r>${run('KẾ TOÁN', '<w:b/>')}<w:r><w:tab/></w:r>${run('GIÁM ĐỐC', '<w:b/>')}`,
        `<w:tabs><w:tab w:val="center" w:pos="4560"/><w:tab w:val="right" w:pos="9000"/></w:tabs>`,
      ),
      // Decimal tab: số tiền thẳng hàng dấu thập phân.
      ...[
        ['Phí khởi tạo', '1.250'],
        ['Thuê bao tháng', '99'],
        ['Phụ phí vùng xa', '12.345'],
      ].map(([label, amount]) =>
        p(
          `${run(label)}<w:r><w:tab/></w:r>${run(amount + ' đ')}`,
          `<w:tabs><w:tab w:val="decimal" w:pos="6000"/></w:tabs>`,
        ),
      ),
    ].join('\n'),
  },
  {
    title: 'Giãn dòng, ngắt dòng mềm & ký hiệu',
    xml: [
      // w:spacing: giãn dòng 1.5 (line=360 auto) + cách trên 12pt / dưới 6pt.
      p(
        run('Đoạn này dùng w:spacing — giãn dòng 1.5, cách đoạn trên 12pt và dưới 6pt nên trông thưa hơn hẳn các đoạn quanh nó. ' + LOREM[2]),
        '<w:spacing w:before="240" w:after="120" w:line="360" w:lineRule="auto"/>',
      ),
      // w:br: ngắt dòng mềm — vẫn cùng một đoạn, không tạo paragraph mới.
      p(
        run('Dòng đầu của đoạn này') +
          '<w:r><w:br/></w:r>' +
          run('được ngắt dòng mềm (w:br) xuống dòng hai mà vẫn nằm trong cùng một paragraph.'),
      ),
      // w:sym: ký tự từ font symbol (Wingdings F0B7 → •) chèn giữa văn bản.
      p(
        '<w:r><w:sym w:font="Wingdings" w:char="F0B7"/></w:r>' +
          run(' Ký hiệu w:sym (Wingdings F0B7 → •) chèn ngay đầu dòng — khác với bullet của danh sách.'),
      ),
    ].join('\n'),
  },
];

function buildDocumentXml() {
  const chapters = SPECIALS.map((s, i) => chapter(i + 1, s));
  // Padding chương cuối để chắc chắn tràn sang trang 3.
  const tail = [
    // w:pageBreakBefore — chương này luôn bắt đầu ở đầu một trang mới.
    p(
      run('9. Phần đệm — bắt đầu trang mới (w:pageBreakBefore)', '<w:b/><w:sz w:val="28"/>'),
      '<w:pageBreakBefore/>' + jc('left'),
    ),
    ...Array.from({ length: 20 }, (_, i) => p(run(LOREM[i % LOREM.length]), jc('both'))),
    heading('10. Bảng thứ hai ở cuối tài liệu'),
    `<w:tbl>
      <w:tblGrid><w:gridCol w:w="4515"/><w:gridCol w:w="4515"/></w:tblGrid>
      <w:tr>${td(p(run('Cột trái', '<w:b/>')))}${td(p(run('Cột phải', '<w:b/>')))}</w:tr>
      <w:tr>${td(p(run('Bảng nằm gần cuối tài liệu để kiểm tra ngắt-trang-nguyên-bảng khi không đủ chỗ.')))}${td(p(run('Ô bên phải.')))}</w:tr>
    </w:tbl>`,
    heading('11. Bảng dài trải qua nhiều trang'),
    p(run('Bảng dưới đây có 36 hàng — dài hơn một trang, buộc layout engine phải ngắt trang theo từng hàng (M5).')),
    `<w:tbl>
      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="6630"/></w:tblGrid>
      <w:tr><w:trPr><w:tblHeader/></w:trPr>${td(p(run('Hàng', '<w:b/>')))}${td(p(run('Mô tả (hàng tiêu đề lặp lại mỗi trang)', '<w:b/>')))}</w:tr>
      ${Array.from({ length: 36 }, (_, i) =>
        `<w:tr>${td(p(run(`Hàng ${i + 1}`)))}${td(p(run(`Nội dung ô bên phải của hàng ${i + 1} — một dòng mô tả đủ dài để chiếm trọn bề ngang cột và đôi khi xuống dòng thứ hai.`)))}</w:tr>`,
      ).join('\n')}
    </w:tbl>`,
    heading('12. Hàng cao hơn một trang'),
    p(run('Hàng thứ nhất của bảng sau chứa 60 đoạn — cao hơn cả một trang, buộc phải tách giữa hàng.')),
    `<w:tbl>
      ${TBL_BORDERS}
      <w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="6630"/></w:tblGrid>
      <w:tr>
        ${td(p(run('Nhãn của hàng siêu cao', '<w:b/>')))}
        ${td(Array.from({ length: 60 }, (_, i) => p(run(`Dòng ${i + 1} trong ô siêu cao — nội dung lấp đầy để tổng chiều cao của hàng vượt quá một trang.`))).join('\n'))}
      </w:tr>
      <w:tr>${td(p(run('Hàng sau')))}${td(p(run('Hàng bình thường ngay sau hàng siêu cao.')))}</w:tr>
    </w:tbl>`,
    p(run('— Hết tài liệu mẫu —', '<w:i/>'), jc('center')),
  ].join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:wp="${WP_NS}" xmlns:a="${A_NS}" xmlns:pic="${PIC_NS}">
  <w:body>
    ${p(run('bapbong sample document', '<w:b/><w:sz w:val="36"/>'), jc('center'))}
    ${p(run('Tài liệu mẫu nhiều trang: đậm/nghiêng/màu, danh sách đa cấp, bảng gộp ô, ảnh inline, hyperlink, canh lề và thụt lề.', '<w:i/>'), jc('center'))}
    ${chapters.join('\n')}
    ${tail}
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rId20"/>
      <w:footerReference w:type="default" r:id="rId21"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

const HEADER_XML = `<?xml version="1.0"?><w:hdr xmlns:w="${W_NS}">
  ${p(run('bapbong sample — tài liệu mẫu nhiều trang', '<w:i/><w:color w:val="6B6B70"/>'), jc('right'))}
</w:hdr>`;

// Footer "Trang X / Y": PAGE qua complex field (fldChar), NUMPAGES qua
// fldSimple — phủ cả hai dạng field của Word.
const GRAY = '<w:i/><w:color w:val="6B6B70"/>';
const FOOTER_XML = `<?xml version="1.0"?><w:ftr xmlns:w="${W_NS}">
  <w:p><w:pPr>${jc('center')}</w:pPr>
    ${run('Trang ', GRAY)}
    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:rPr>${GRAY}</w:rPr><w:t>1</w:t></w:r>
    <w:r><w:fldChar w:fldCharType="end"/></w:r>
    ${run(' / ', GRAY)}
    <w:fldSimple w:instr=" NUMPAGES "><w:r><w:rPr>${GRAY}</w:rPr><w:t>7</w:t></w:r></w:fldSimple>
    ${run(' — bapbong playground', GRAY)}
  </w:p>
</w:ftr>`;

const STYLES_XML = `<?xml version="1.0"?><w:styles xmlns:w="${W_NS}">
  <w:style w:styleId="Indented"><w:pPr><w:ind w:left="720" w:right="720"/></w:pPr></w:style>
  <w:style w:styleId="Quote"><w:basedOn w:val="Indented"/><w:rPr><w:i/><w:color w:val="555555"/></w:rPr></w:style>
</w:styles>`;

const NUMBERING_XML = `<?xml version="1.0"?><w:numbering xmlns:w="${W_NS}">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:start w:val="1"/></w:lvl>
    <w:lvl w:ilvl="1"><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2."/><w:start w:val="1"/></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

const RELS_XML = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}">
  <Relationship Id="rId7" Type="${R_NS}/image" Target="media/image1.png"/>
  <Relationship Id="rId8" Type="${R_NS}/image" Target="media/image2.png"/>
  <Relationship Id="rId9" Type="${R_NS}/hyperlink" Target="https://prosemirror.net/" TargetMode="External"/>
  <Relationship Id="rId20" Type="${R_NS}/header" Target="header1.xml"/>
  <Relationship Id="rId21" Type="${R_NS}/footer" Target="footer1.xml"/>
</Relationships>`;

async function main() {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${R_NS}/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file('word/document.xml', buildDocumentXml());
  zip.file('word/styles.xml', STYLES_XML);
  zip.file('word/numbering.xml', NUMBERING_XML);
  zip.file('word/_rels/document.xml.rels', RELS_XML);
  zip.file('word/header1.xml', HEADER_XML);
  zip.file('word/footer1.xml', FOOTER_XML);
  zip.file('word/media/image1.png', PNG_RED, { base64: true });
  zip.file('word/media/image2.png', PNG_BLUE, { base64: true });

  const bytes = await zip.generateAsync({ type: 'nodebuffer' });
  const out = path.join(__dirname, '..', 'public', 'sample.docx');
  fs.writeFileSync(out, bytes);
  console.log(`wrote ${out} (${bytes.length} bytes)`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
