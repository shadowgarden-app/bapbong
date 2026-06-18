// Generates apps/playground/public/sample.docx — a multi-page document that
// exercises the import + layout + paint pipeline end to end: marks (bold,
// italic, color, size, underline, strike, highlight, super/subscript),
// alignment (center/right/justify), first-line and hanging indent, paragraph
// spacing (w:spacing) + soft line breaks (w:br) + symbol runs (w:sym) +
// page-break-before, multi-level numbered + bullet lists, tables (colwidth,
// colspan, rowspan/vMerge, table align, row height, cell vAlign, per-cell &
// table borders, shading), inline + floating images, hyperlinks, tab stops,
// footnotes (laid out at the bottom of the page their reference falls on), a
// two-column section (w:cols) introduced by a continuous section break,
// first/even header variants (w:titlePg + w:evenAndOddHeaders), and comments
// (w:commentRange + comments.xml) shown in a sidebar.
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

// A footnote reference run — renders as a superscript number; the body lives
// in footnotes.xml and is laid out at the bottom of the page it lands on.
const fnRef = (id) => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;

// Wrap run content in a comment range (the body lives in comments.xml).
const commented = (id, content) =>
  `<w:commentRangeStart w:id="${id}"/>${content}<w:commentRangeEnd w:id="${id}"/><w:r><w:commentReference w:id="${id}"/></w:r>`;

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
        commented(2, run('đậm nghiêng', '<w:b/><w:i/>')),
        run(', '),
        commented(0, run('màu đỏ', '<w:color w:val="C0392B"/>')),
        run(', '),
        commented(3, run('màu xanh', '<w:color w:val="1F6FEB"/>')),
        run(', '),
        commented(4, run('gạch chân', '<w:u w:val="single"/>')),
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
        run(' trong cùng một dòng để kiểm tra đo chữ theo từng run'),
        fnRef(1),
        run('.'),
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
        ${td(p(run('158 đơn hàng, tăng trưởng tốt nhờ chiến dịch mới') + fnRef(3)))}
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
        run('Đoạn này dùng w:spacing — giãn dòng 1.5, cách đoạn trên 12pt và dưới 6pt nên trông thưa hơn hẳn các đoạn quanh nó') +
          fnRef(2) +
          run('. ' + LOREM[2]),
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
    heading('13. Bố cục nhiều cột (w:cols)'),
    // This paragraph carries the section break that CLOSES the single-column
    // section (everything above). The next section flows in two columns.
    p(
      run('Đoạn giới thiệu trình bày một cột. ') +
        commented(1, run('Phần thân ngay bên dưới chuyển sang HAI cột bằng một section break liên tục')) +
        run(' (w:type=continuous): văn bản rót đầy cột trái rồi mới sang cột phải, đúng kiểu bản tin / tạp chí.'),
      jc('both') + '<w:sectPr><w:type w:val="continuous"/><w:cols w:num="1"/></w:sectPr>',
    ),
    // Two-column body — the LAST paragraph carries the 2-column section break.
    ...Array.from({ length: 7 }, (_, i) =>
      p(run(`(${i + 1}) ` + LOREM[i % LOREM.length]), jc('both')),
    ),
    p(
      run('(8) ' + LOREM[3] + ' ' + LOREM[5]),
      jc('both') + '<w:sectPr><w:type w:val="continuous"/><w:cols w:num="2" w:space="425"/></w:sectPr>',
    ),
    p(run('Trở lại bố cục một cột sau section break thứ hai — phần kết của tài liệu chiếm trọn bề ngang trang trở lại.'), jc('both')),
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
      <w:headerReference w:type="first" r:id="rId23"/>
      <w:headerReference w:type="even" r:id="rId24"/>
      <w:titlePg/>
      <w:type w:val="continuous"/>
      <w:cols w:num="1"/>
    </w:sectPr>
  </w:body>
</w:document>`;
}

// Default header = odd pages (right-aligned). With w:titlePg + w:evenAndOddHeaders,
// page 1 uses HEADER_FIRST_XML and even pages use HEADER_EVEN_XML.
const HEADER_XML = `<?xml version="1.0"?><w:hdr xmlns:w="${W_NS}">
  ${p(run('bapbong sample — trang lẻ', '<w:i/><w:color w:val="6B6B70"/>'), jc('right'))}
</w:hdr>`;
const HEADER_FIRST_XML = `<?xml version="1.0"?><w:hdr xmlns:w="${W_NS}">
  ${p(run('« TRANG BÌA — header riêng cho trang đầu »', '<w:b/><w:color w:val="1F6FEB"/>'), jc('center'))}
</w:hdr>`;
const HEADER_EVEN_XML = `<?xml version="1.0"?><w:hdr xmlns:w="${W_NS}">
  ${p(run('trang chẵn — bapbong sample', '<w:i/><w:color w:val="6B6B70"/>'), jc('left'))}
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

// Footnote bodies: the separator notes (negative / 0 ids) are skipped on
// import; ids 1 and 2 are the real notes referenced from the body.
const FOOTNOTES_XML = `<?xml version="1.0"?><w:footnotes xmlns:w="${W_NS}">
  <w:footnote w:id="-1" w:type="separator"><w:p><w:r><w:separator/></w:r></w:p></w:footnote>
  <w:footnote w:id="0" w:type="continuationSeparator"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>
  <w:footnote w:id="1"><w:p>${run('Chú thích chân trang đầu tiên — bản thân ghi chú được layout ở đáy của chính trang chứa tham chiếu, với cỡ chữ nhỏ hơn phần thân.')}</w:p></w:footnote>
  <w:footnote w:id="2"><w:p>${run('Chú thích thứ hai nằm ở một trang khác, chứng minh mỗi trang tự gom ghi chú riêng phía trên footer.')}</w:p></w:footnote>
  <w:footnote w:id="3"><w:p>${run('Chú thích này được tham chiếu từ bên trong một ô bảng — vẫn được gom xuống đáy trang chứa bảng.')}</w:p></w:footnote>
</w:footnotes>`;

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

// Document settings — w:evenAndOddHeaders turns on the even-page header variant.
const SETTINGS_XML = `<?xml version="1.0"?><w:settings xmlns:w="${W_NS}"><w:evenAndOddHeaders/></w:settings>`;

// Comment bodies referenced by w:commentRange in the document body.
// ids 0/2/3/4 land in the same "Chữ đậm…" paragraph, so their anchors cluster
// on nearby lines — a ready-made case for the comment collision-avoidance /
// active-snap behaviour (minimize bubbles + expand cards). id 1 sits far away
// (multi-column section) as the non-colliding control.
const COMMENTS_XML = `<?xml version="1.0"?><w:comments xmlns:w="${W_NS}">
  <w:comment w:id="0" w:author="Phương Minh" w:date="2026-06-17T09:00:00Z" w:initials="PM"><w:p><w:r><w:t>Cụm "màu đỏ" được tô comment — bấm vào đây để chọn vùng tương ứng trên canvas.</w:t></w:r></w:p></w:comment>
  <w:comment w:id="1" w:author="Reviewer" w:date="2026-06-17T10:30:00Z" w:initials="RV"><w:p><w:r><w:t>Comment vắt qua nhiều run trong phần bố cục nhiều cột.</w:t></w:r></w:p></w:comment>
  <w:comment w:id="2" w:author="Minh Lê" w:date="2026-06-17T09:05:00Z" w:initials="ML"><w:p><w:r><w:t>Comment trên "đậm nghiêng" — neo sát các comment kế bên để test chống đè.</w:t></w:r></w:p></w:comment>
  <w:comment w:id="3" w:author="Reviewer" w:date="2026-06-17T09:10:00Z" w:initials="RV"><w:p><w:r><w:t>Comment trên "màu xanh" cùng dòng.</w:t></w:r></w:p></w:comment>
  <w:comment w:id="4" w:author="Phương Minh" w:date="2026-06-17T09:15:00Z" w:initials="PM"><w:p><w:r><w:t>Comment trên "gạch chân" — cụm này dùng để kiểm tra active-snap khi chọn từng card/bong bóng.</w:t></w:r></w:p></w:comment>
</w:comments>`;

const RELS_XML = `<?xml version="1.0"?><Relationships xmlns="${PKG_REL_NS}">
  <Relationship Id="rId7" Type="${R_NS}/image" Target="media/image1.png"/>
  <Relationship Id="rId8" Type="${R_NS}/image" Target="media/image2.png"/>
  <Relationship Id="rId9" Type="${R_NS}/hyperlink" Target="https://prosemirror.net/" TargetMode="External"/>
  <Relationship Id="rId20" Type="${R_NS}/header" Target="header1.xml"/>
  <Relationship Id="rId21" Type="${R_NS}/footer" Target="footer1.xml"/>
  <Relationship Id="rId22" Type="${R_NS}/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId23" Type="${R_NS}/header" Target="header2.xml"/>
  <Relationship Id="rId24" Type="${R_NS}/header" Target="header3.xml"/>
  <Relationship Id="rId25" Type="${R_NS}/settings" Target="settings.xml"/>
  <Relationship Id="rId26" Type="${R_NS}/comments" Target="comments.xml"/>
</Relationships>`;

async function main() {
  const zip = new JSZip();
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`,
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
  zip.file('word/header2.xml', HEADER_FIRST_XML);
  zip.file('word/header3.xml', HEADER_EVEN_XML);
  zip.file('word/footer1.xml', FOOTER_XML);
  zip.file('word/footnotes.xml', FOOTNOTES_XML);
  zip.file('word/settings.xml', SETTINGS_XML);
  zip.file('word/comments.xml', COMMENTS_XML);
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
