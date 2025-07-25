function extractArticle(row) {
  const pattern = /(KR|KU|КР|КУ|KLT|РТ|PT)[-–]?(\d+)(?:[-–.]?(\d+))?/i;

  for (let cell of row) {
    const match = typeof cell === 'string' && cell.match(pattern);
    if (match) {
      const prefix = match[1].toUpperCase();

      // 🎯 Особый случай: если префикс PT → озвучиваем всю строку
      if (prefix === "PT") {
        return row.filter(Boolean).join(", ");
      }

      // Стандартная озвучка по префиксам
      return formatArticle(match[1], match[2], match[3]);
    }
  }

  return null;
}
