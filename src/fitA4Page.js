// Giữ mỗi trang in trong khổ A4: trang nào tràn thì giảm cỡ chữ bảng cho vừa.
export const MAX_ROWS_PER_PAGE = 54;
const MIN_FONT_SIZE_PT = 8;
const FONT_STEP_PT = 0.25;

export function fitFontSize(startSize, overflowsAt) {
  let size = startSize;
  while (size > MIN_FONT_SIZE_PT && overflowsAt(size)) size -= FONT_STEP_PT;
  return size;
}

function pageOverflows(page, table) {
  const contentLimit = page.getBoundingClientRect().bottom - parseFloat(getComputedStyle(page).paddingBottom);
  const contentBottom = Math.max(table.getBoundingClientRect().bottom, page.lastElementChild.getBoundingClientRect().bottom);
  return contentBottom > contentLimit + 0.5;
}

export function fitA4Pages(pages, tableSize) {
  pages.forEach((page) => {
    const table = page.querySelector('table');
    if (!table) return;
    fitFontSize(tableSize, (size) => {
      table.style.fontSize = `${size}pt`;
      return pageOverflows(page, table);
    });
  });
}
