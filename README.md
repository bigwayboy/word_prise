# 自助格式报价单

这是一个纯前端网页，客户上传 `.docx` 后会在浏览器本地分析文档并计算报价。论文文件不会上传到服务器。

## 改价格

打开 `app.js` 顶部的 `PRICE_CONFIG` 即可修改价格：

```js
const PRICE_CONFIG = {
  textFormatPerThousand: 4,
  toc: 15,
  cover: 10,
  headerFooter: 15,
  figuresTables: 0.9,
  references: 1.2,
  formulas: 0.9,
  footnotes: 0.8,
  rushRate: 0.3,
  freeCoverThreshold: 100,
};
```

## 发给客户使用

推荐把 `index.html`、`styles.css`、`app.js` 三个文件上传到静态网站托管平台，然后发客户一个网页链接。

可选方式：

- GitHub Pages：免费，适合长期放一个固定报价页。
- Netlify / Vercel：拖拽整个文件夹即可发布，操作更简单。
- 自己的网站空间：把三个文件放到同一个目录，让客户访问 `index.html`。

不建议直接把文件夹发给客户，因为不同电脑的浏览器安全策略不同，直接双击打开可能影响 `.docx` 解压能力。使用 HTTPS 网页链接最稳定。
