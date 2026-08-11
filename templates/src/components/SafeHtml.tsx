import React, { useMemo } from 'react';

interface SafeHtmlProps {
  html: string;
  className?: string;
  minHeight?: number;
}

/**
 * 用沙箱 iframe 渲染邮件 HTML，默认禁止脚本与同源访问。
 */
export function SafeHtml({ html, className, minHeight = 280 }: SafeHtmlProps) {
  const srcDoc = useMemo(() => {
    const content = html || '';
    return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <base target="_blank" rel="noopener noreferrer" />
    <style>
      :root { color-scheme: light dark; }
      body {
        margin: 0;
        padding: 12px;
        font: 14px/1.55 -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        color: #0f172a;
        word-break: break-word;
        background: transparent;
      }
      @media (prefers-color-scheme: dark) {
        body { color: #e2e8f0; }
      }
      img, video, iframe { max-width: 100%; height: auto; }
      a { color: #2563eb; }
      pre, code { white-space: pre-wrap; word-break: break-word; }
      table { max-width: 100%; border-collapse: collapse; }
    </style>
  </head>
  <body>${content}</body>
</html>`;
  }, [html]);

  return (
    <iframe
      title="email-html-body"
      sandbox=""
      referrerPolicy="no-referrer"
      srcDoc={srcDoc}
      className={className}
      style={{ width: '100%', minHeight, border: 0, borderRadius: 8, background: 'transparent' }}
    />
  );
}

export default SafeHtml;
