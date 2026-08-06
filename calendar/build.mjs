#!/usr/bin/env node
// ショートステイ管理表のビルドスクリプト。
//
// 配布物は「1ファイルで完結するHTML」である必要がある（施設のPCにネット接続が
// ない場合でも開けるようにするため）。フォントとReactランタイムはbase bundleに
// gzip+base64で同梱済みなので、それらはそのまま流用し、
// テンプレート(src/template.html)とロジック(src/logic.js)だけを差し替える。
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

const here = dirname(fileURLToPath(import.meta.url));
const OUT = join(here, 'dist', 'ショートステイ管理表_v5.html');

// フォントとReactの同梱物(約8MB)はビルドしても中身が変わらない。同じものを2つ
// リポジトリに置く意味はないので、元のv4があればそれを、無ければ前回のビルド成果物
// を素材元として使う。どちらもマニフェストは同一。
const ORIGINAL = join(here, 'assets', 'base-bundle-v4.html');
const BASE = existsSync(ORIGINAL) ? ORIGINAL : OUT;
if (!existsSync(BASE)) throw new Error('素材元が見つかりません: ' + ORIGINAL + ' か ' + OUT + ' が必要です');

// bundleは1行=1ブロックの構造。行番号ではなく「script type」で位置を特定する。
const bundle = readFileSync(BASE, 'utf8');
const lines = bundle.split('\n');

const templateTagIdx = lines.findIndex((l) => l.includes('<script type="__bundler/template">'));
if (templateTagIdx < 0) throw new Error('base bundle に __bundler/template が見つかりません');
const pageLineIdx = templateTagIdx + 1;

const originalPage = JSON.parse(lines[pageLineIdx]);

// helmet 内の最初の <style> が @font-face 定義（約7,000行相当）。
// ここを引き抜いて新テンプレートへ再注入する。
const fontMatch = originalPage.match(/<style>@font-face[\s\S]*?<\/style>/);
if (!fontMatch) throw new Error('base bundle からフォント定義を抽出できません');
const fontCss = fontMatch[0];

// <head> はランタイムscriptのuuid参照を含むのでそのまま使う。
const headMatch = originalPage.match(/^([\s\S]*?<body>)/);
if (!headMatch) throw new Error('base bundle から <head> を抽出できません');
const headHtml = headMatch[1];

// data-props はホスト側のプロパティエディタ定義。既存の値を引き継ぐ。
const propsMatch = originalPage.match(/<script type="text\/x-dc" data-dc-script="" data-props="([^"]*)">/);
const dataProps = propsMatch ? propsMatch[1] : '';

const template = readFileSync(join(here, 'src', 'template.html'), 'utf8')
  .replace('<!--FONT_FACE_CSS-->', fontCss);
const logic = readFileSync(join(here, 'src', 'logic.js'), 'utf8');

const page = [
  headHtml,
  '<x-dc>',
  template,
  '</x-dc>',
  `<script type="text/x-dc" data-dc-script="" data-props="${dataProps}">`,
  logic,
  '</script>',
  '</body></html>',
].join('\n');

// この JSON は <script type="__bundler/template"> の中身として HTML に埋め込まれる。
// ページ自身が </script> を含む（x-dc のロジックスクリプト）ため、素の JSON だと
// そこで外側の script タグが閉じてしまう。'<' を \u003c へ逃がして防ぐ
// （元のバンドルも同じ方法を取っている）。
lines[pageLineIdx] = JSON.stringify(page).replace(/</g, '\\u003c').replace(/>/g, '\\u003e');
writeFileSync(OUT, lines.join('\n'), 'utf8');

const kb = (n) => (n / 1024).toFixed(0) + 'KB';
console.log(`built ${OUT}`);
console.log(`  page: ${kb(page.length)} (font-face ${kb(fontCss.length)} / template ${kb(template.length)} / logic ${kb(logic.length)})`);
