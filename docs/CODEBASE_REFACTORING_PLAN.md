# Codebase Refactoring Plan

この計画は、描画精度改善の前提としてコード配置と責務境界を作り直すためのもの。後方互換性は前提にしない。既存 API を温存するための複雑化も避ける。

## 目的

- 1ファイル 800行以下を原則にする。
- parser、document model、layout、renderer、asset resolver、test helper を明確に分ける。
- `html-renderer-sync.ts` と `document-parser.ts` の巨大 class を解体する。
- section / page break の本格実装を入れられる構造にする。
- lint を通すためだけの読みにくい抽象化は作らない。

## 現状の主な問題

### 1. 巨大ファイルに責務が集中している

現状の最大ファイル:

- `src/html-renderer-sync.ts`: 3349 lines
- `src/document-parser.ts`: 2820 lines

この2ファイルが、parse、state mutation、pagination、DOM creation、overflow measurement、header/footer rendering、shape rendering、math rendering、table rendering を抱えている。結果として、section/page break のような横断機能を入れると、ほぼ全域に影響が出る。

### 2. Parser と Renderer の中間層が弱い

XML parse 後の model はあるが、layout に必要な意味づけが renderer 側へ流れ込んでいる。

例:

- paragraph-level `sectPr` から synthetic `SectionBreak` を renderer が挿入している。
- section type の意味を renderer が後ろから補正している。
- `Page` が layout result なのか render work item なのか曖昧。
- overflow split のために model object に `breakIndex` などの render-time state が混ざる。

### 3. DOM rendering と pagination が密結合している

`renderElements` が DOM を append しながら overflow を検出し、その場で page を増やす。これは小さい文書では動くが、次の変更が難しくなる。

- region-based section layout。
- header/footer 測定後の再 pagination。
- column flow の制御。
- table / paragraph continuation の独立テスト。

### 4. Feature renderer の境界がない

paragraph、run、table、drawing、shape、math、header/footer、notes が同じ class 内で処理されている。1つの feature を直すために unrelated code を読み続ける必要がある。

### 5. Test が構造変更を支えにくい

browser smoke はあるが、parser output、section stream、pagination decision、DOM renderer unit を個別に固定するテストが少ない。巨大 class を分割する前に、最低限の characterization test が必要。

## 新しい責務境界

### Public Entry

外部入力と全体 orchestration だけを持つ。

Suggested files:

- `src/index.ts`
- `src/render.ts`
- `src/options.ts`

責務:

- input normalization。
- package loading。
- parse pipeline 実行。
- layout pipeline 実行。
- DOM renderer 実行。
- result object 作成。

### Package Layer

DOCX package、part、relationship、mime を扱う。

Suggested files:

- `src/package/open-xml-package.ts`
- `src/package/part.ts`
- `src/package/relationship.ts`
- `src/package/content-types.ts`

既存の `src/common` と `src/mime` はここへ寄せる。

### Parse Layer

XML から document model を作る。layout 判断と DOM rendering はしない。

Suggested files:

- `src/parser/document-parser.ts`
- `src/parser/body-parser.ts`
- `src/parser/paragraph-parser.ts`
- `src/parser/run-parser.ts`
- `src/parser/table-parser.ts`
- `src/parser/drawing-parser.ts`
- `src/parser/math-parser.ts`
- `src/parser/section-parser.ts`
- `src/parser/header-footer-parser.ts`
- `src/parser/style-parser.ts`
- `src/parser/numbering-parser.ts`

移行元:

- `src/document-parser.ts`
- `src/document/paragraph.ts`
- `src/document/run.ts`
- `src/document/section.ts`
- feature-specific parse code。

### Model Layer

Parse 結果の型だけを持つ。DOM や browser API に依存しない。

Suggested files:

- `src/model/document.ts`
- `src/model/block.ts`
- `src/model/inline.ts`
- `src/model/table.ts`
- `src/model/drawing.ts`
- `src/model/math.ts`
- `src/model/section.ts`
- `src/model/page.ts`
- `src/model/style.ts`

ルール:

- model に `HTMLElement` を入れない。
- model に render-time mutation state を入れない。
- `breakIndex` のような split state は layout layer の working data に限定する。

### Layout Layer

section stream、pagination、column region、header/footer context、overflow decision を扱う。

Suggested files:

- `src/layout/section-stream.ts`
- `src/layout/layout-region.ts`
- `src/layout/page-builder.ts`
- `src/layout/pagination.ts`
- `src/layout/breaks.ts`
- `src/layout/overflow.ts`
- `src/layout/header-footer-context.ts`
- `src/layout/page-numbering.ts`

責務:

- parsed blocks を section region stream に変換する。
- explicit break を data として解釈する。
- physical page と layout region を作る。
- overflow measurement の結果を split decision に変換する。
- header/footer 解決に必要な page context を作る。

### Render Layer

Layout result を DOM へ変換する。pagination policy は持たない。

Suggested files:

- `src/render/dom-renderer.ts`
- `src/render/styles/default-styles.ts`
- `src/render/styles/document-styles.ts`
- `src/render/page-renderer.ts`
- `src/render/region-renderer.ts`
- `src/render/block-renderer.ts`
- `src/render/inline-renderer.ts`
- `src/render/table-renderer.ts`
- `src/render/drawing-renderer.ts`
- `src/render/math-renderer.ts`
- `src/render/header-footer-renderer.ts`
- `src/render/notes-renderer.ts`
- `src/render/fields-renderer.ts`

ルール:

- render layer は DOM を作るだけ。
- page split を決めない。
- layout layer から渡された `PhysicalPage` / `LayoutRegion` を素直に描画する。

### Measurement Layer

Browser DOM measurement を隔離する。

Suggested files:

- `src/measure/dom-measurer.ts`
- `src/measure/overflow-measurer.ts`
- `src/measure/header-footer-measurer.ts`

責務:

- `clientHeight` / `scrollHeight` 比較。
- hidden measurement container。
- header/footer height measurement。
- layout layer に返す measurement result の整形。

### Assets Layer

image、font、theme、relationship resolution を扱う。

Suggested files:

- `src/assets/image-resolver.ts`
- `src/assets/font-resolver.ts`
- `src/assets/theme-resolver.ts`
- `src/assets/relationship-resolver.ts`

### Compatibility Layer

後方互換は不要だが、移行中に一時 adapter は置いてよい。

Suggested files:

- `src/legacy/html-renderer-adapter.ts`
- `src/legacy/document-parser-adapter.ts`

ルール:

- legacy adapter は移行完了後に削除する。
- adapter 内だけ 800行制限を一時的に超えることも許す。ただし最終状態では残さない。

## 800行ルール

最終状態:

- Production `.ts` file は 800行以下。
- 目標は 300-500行程度。
- 800行を超える場合は責務境界が間違っている可能性が高いので分割する。
- Generated file や golden fixture は対象外。

分割判断:

- import が feature をまたぎすぎたら分割する。
- class private method が20個を超えたら分割を検討する。
- DOM API と XML API が同じ file に出たら原則分ける。
- layout decision と DOM append が同じ function に出たら分ける。

## Target Directory Shape

```text
src/
  assets/
  layout/
  measure/
  model/
  package/
  parser/
  render/
    styles/
  shared/
  test-utils/
```

既存 directory は段階的に移す。

```text
src/document/          -> src/model/ and src/parser/
src/common/            -> src/package/ or src/shared/
src/header-footer/     -> src/model/, src/parser/, src/render/
src/notes/             -> src/model/, src/parser/, src/render/
src/theme/             -> src/assets/ or src/model/
src/styles/            -> src/parser/ and src/render/styles/
```

## Refactoring Phases

### Phase 0: Characterization Tests

大きく動かす前に、現状を固定する。

- Browser smoke は維持。
- Parser output の unit test を追加する。
- section stream の unit test を追加する。
- header/footer reference inheritance の unit test を追加する。
- simple document の DOM structure snapshot を追加する。

この phase では production code を大きく変えない。

### Phase 1: Model Types を分離する

`src/document/dom.ts` と周辺 type を `src/model` へ移す。

やること:

- DOM に依存しない type を `src/model/*` に分ける。
- `Page` から render-time field を剥がす準備をする。
- import path を整理する。

やらないこと:

- pagination behavior は変えない。
- renderer logic は変えない。

### Phase 2: Parser を feature ごとに分割する

`src/document-parser.ts` を 800行以下の parser modules に割る。

分割順:

1. section parser。
2. paragraph / run parser。
3. table parser。
4. drawing / image / shape parser。
5. math parser。
6. style / numbering / header-footer parser。

完了条件:

- `src/document-parser.ts` は orchestration だけになり 300行以下。
- 各 parser module は DOM rendering import を持たない。

### Phase 3: Renderer を feature renderer に分割する

`src/html-renderer-sync.ts` から DOM rendering を feature ごとに切り出す。

分割順:

1. default/document style rendering。
2. page/header/footer rendering。
3. paragraph/run/text rendering。
4. table rendering。
5. drawing/image/shape rendering。
6. math rendering。
7. notes/fields/revisions。

完了条件:

- `html-renderer-sync.ts` は orchestration adapter だけになり 500行以下。
- feature renderer は layout decision を持たない。

### Phase 4: Layout Layer を導入する

section/page break 改善計画と接続する。

やること:

- `LayoutRegion` と `PhysicalPage` を導入する。
- section stream を renderer ではなく layout layer で作る。
- explicit break を layout data として解釈する。
- page context を header/footer resolver に渡す。

完了条件:

- 1 physical page に複数 region を持てる。
- `Page` に単一 `sectProps` を持たせる設計から脱却する。

### Phase 5: Measurement と Pagination を分離する

DOM measurement と split decision を分ける。

やること:

- `checkOverflow` を `measure` layer へ移す。
- append 中の即時 page mutation をやめる。
- measured result から layout layer が next page / next region を決める。
- `lastRenderedPageBreak` を hint として扱う。

完了条件:

- `renderElements` は page 配列を直接 splice しない。
- split decision は unit test できる pure-ish module に寄る。

### Phase 6: Public API を整理する

後方互換不要なので、使いやすい entry に作り直す。

案:

```ts
const result = await renderDocx(source, {
  pagination: "page",
  assets: { imageUrlMode: "object-url" },
});
```

旧 API adapter はこの時点で削除する。

### Phase 7: Cleanup

- legacy adapter を削除する。
- 使われなくなった `src/document` / `src/common` を消す。
- line count check を追加する。
- docs を新 architecture に更新する。

## Line Count Check

CI または local test に、production file の line count check を追加する。

対象:

- `src/**/*.ts`

除外:

- `src/typings.d.ts`
- generated file が将来できた場合の generated directory。

失敗条件:

- 800行超過。

この check は最後に追加する。移行途中に先に入れると作業を邪魔する。

## Migration Rules

- Behavior change と file move を同じ PR に詰め込まない。
- まず characterization test、その後 move、最後に behavior change。
- Parser split 中は layout/rendering を変えない。
- Renderer split 中は parse output を変えない。
- Layout rewrite は section/page break 専用 test と一緒に進める。
- 旧 code を残す場合は `legacy` と明示し、削除 phase を決める。

## What Not To Do

- `utils.ts` に何でも置かない。
- `renderer-helper.ts` のような曖昧な巨大 file を作らない。
- 互換維持のための alias と overload を増やさない。
- parser から DOM node を作らない。
- renderer で XML element を直接読む設計に戻さない。
- layout decision を CSS hack だけで済ませない。

## First Implementation Slice

最初の実装範囲は小さくする。

1. `src/model` を作る。
2. `src/document/dom.ts` の型を feature ごとに model file へ分割する。
3. import path を更新する。
4. browser smoke と unit test を通す。
5. `src/document-parser.ts` と `src/html-renderer-sync.ts` の中身はまだ大きく変えない。

これで behavior を変えずに、以降の分割作業の土台を作る。

## Connection To Section/Page Break Work

section/page break の本格修正は、この refactor の Phase 4 以降で行うのが安全。特に以下は refactor なしに入れるべきではない。

- 1 page multiple regions。
- section-local page index。
- header/footer 測定後の re-pagination。
- `lastRenderedPageBreak` の hint 化。

先に parser/model/render の境界を作ることで、描画精度改善の変更範囲を小さくできる。
