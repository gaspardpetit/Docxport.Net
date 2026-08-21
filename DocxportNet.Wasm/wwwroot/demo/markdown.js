const escapeHtml = value => value.replace(/[&<>"']/g, char => ({
  "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#39;"
}[char]));

function inline(value) {
  const htmlTokens = [];
  const protectedValue = value.replace(
    /<!--[\s\S]*?-->|<\/?[A-Za-z][^>]*>|&(?:#\d+|#x[\dA-Fa-f]+|[A-Za-z][A-Za-z\d]+);/g,
    match => {
    const token = `\u0000HTML${htmlTokens.length}\u0000`;
    htmlTokens.push(match);
    return token;
    });

  let rendered = escapeHtml(protectedValue)
    .replace(/!\[([^\]]*)\]\(([^ )]+)(?:\s+"[^"]*")?\)/g, '<img alt="$1" src="$2">')
    .replace(/\[([^\]]+)\]\(([^ )]+)(?:\s+"[^"]*")?\)/g, '<a href="$2" target="_blank" rel="noreferrer">$1</a>')
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/__([^_]+)__/g, '<strong>$1</strong>')
    .replace(/\*([^*]+)\*/g, '<em>$1</em>');

  rendered = rendered.replace(/\u0000HTML(\d+)\u0000/g, (_, index) => htmlTokens[Number(index)]);
  return rendered;
}

export function renderMarkdown(source) {
  const lines = source.replace(/\r\n?/g, "\n").split("\n");
  const html = [];
  let inCode = false;
  let list = null;
  let paragraph = [];
  const flushParagraph = () => {
    if (paragraph.length) html.push(`<p>${inline(paragraph.join(" "))}</p>`);
    paragraph = [];
  };
  const closeList = () => {
    if (list) html.push(`</${list}>`);
    list = null;
  };

  for (const line of lines) {
    if (/^```/.test(line)) {
      flushParagraph(); closeList();
      html.push(inCode ? "</code></pre>" : "<pre><code>");
      inCode = !inCode;
      continue;
    }
    if (inCode) { html.push(`${escapeHtml(line)}\n`); continue; }
    if (/^\s*(?:<!--|<\/?[A-Za-z][^>]*>)/.test(line)) {
      flushParagraph(); closeList(); html.push(inline(line));
      continue;
    }
    const heading = /^(#{1,6})\s+(.*)$/.exec(line);
    if (heading) {
      flushParagraph(); closeList();
      html.push(`<h${heading[1].length}>${inline(heading[2])}</h${heading[1].length}>`);
      continue;
    }
    const item = /^\s*([-*+] |\d+\. )(.*)$/.exec(line);
    if (item) {
      flushParagraph();
      const kind = /\d/.test(item[1]) ? "ol" : "ul";
      if (list !== kind) { closeList(); html.push(`<${kind}>`); list = kind; }
      html.push(`<li>${inline(item[2])}</li>`);
      continue;
    }
    if (!line.trim()) { flushParagraph(); closeList(); continue; }
    if (/^>\s?/.test(line)) {
      flushParagraph(); closeList(); html.push(`<blockquote>${inline(line.replace(/^>\s?/, ""))}</blockquote>`);
      continue;
    }
    paragraph.push(line.trim());
  }
  flushParagraph(); closeList();
  if (inCode) html.push("</code></pre>");
  return html.join("\n");
}
