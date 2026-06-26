# セクション・改ページ描画精度 改善計画

この計画は、Word のセクション、改ページ、段組み、ヘッダー、フッター、ページネーションの描画精度を上げるためのもの。局所パッチを増やすのではなく、現在の弱い前提を置き換えて、Word の構造を表現できるレイアウトモデルへ寄せる。

## 対象範囲

対象:

- `w:sectPr` のセクション境界と section type。
- 明示的な page break / column break。
- `w:lastRenderedPageBreak` の扱い。
- 複数カラム section。
- section と page をまたぐ header / footer 解決。
- overflow 検出とページ分割。

対象外:

- 図形ジオメトリ精度。
- 改ページ・段組みとの関係を除く OMML 数式スタイル。
- ページ分割に関係しない通常の table layout 精度。

## 現状の流れ

現在のレンダリングフロー:

1. `DocumentParser` が body children を parse し、body 直下の最終 `sectPr` を document root に保持する。
2. paragraph 内の `sectPr` は paragraph props に入る。
3. `HtmlRendererSync.splitPageBySymbol` が synthetic `SectionBreak` node を挿入し、`Page` 配列を作る。
4. 各 `Page` は `sectProps` を1つだけ持つ。
5. `renderPage` が物理ページに相当する `section.docx` と、その中の `article` を1つ作る。
6. `createPageContent` が page 全体の `article` に column CSS を付ける。
7. overflow は `article.clientHeight` と `article.scrollHeight` の比較で検出する。

重要な制約:

- `Page` は `SectionProperties` を1つしか持てない。
- `Continuous` section break は新規 page を作らないが、その地点から新しい margin / page size / columns を適用できない。
- `NextColumn` section break は CSS の `break-before: column` に依存しているが、以降の content を新しい section layout region に束ねられない。
- `createPageContent` には「1ページ内に複数 section があり、section ごとに column が違う」ケースの TODO が残っている。
- header / footer 解決はあるが、page 配列 index と `isFirstPage` に依存しており、overflow で page が増えると壊れやすい。
- `lastRenderedPageBreak` は Word の過去の layout snapshot であり、信頼できる layout rule ではない。

## 主要な問題

### 1. Page と Section が同じ概念として扱われている

Word の section は論理的な書式範囲で、page は物理的な出力結果。現在は `Page` に `sectProps` を直接持たせているため、1 page に 1 section layout しか表現できない。

壊れるケース:

- continuous section が page の途中から始まる。
- 1 page 内で 1カラム content の後に 2カラム content が続く。
- section break で margin や header/footer reference が変わるが、物理 page はまだ変わらない。

これは最優先で直すべき構造問題。`Continuous` や column の局所修正だけでは限界が残る。

### 2. Continuous section が実質的に layout へ反映されない

`Continuous` で新規 page を作らない判断自体は正しい。ただし DOM が 1 page = 1 article なので、section break 以降の layout を切り替えられない。

必要な挙動:

- 同じ物理 page 上で続行する。
- section break 以降に新しい layout region を開始する。
- 新 region に、新 section の columns、docGrid、必要な section state を適用する。

### 3. NextColumn が CSS hint に留まっている

`NextColumn` は section break element に `break-before: column` を付けるだけでは足りない。次の content がどの section layout に属するかを表現できないため、section と column の切り替えが同時に起きる文書で崩れる。

必要な挙動:

- 通常の column break は現在 section の column flow 内で次 column へ進む。
- `NextColumn` section break は次 column へ進み、以降の active section も切り替える。
- 次 column が存在しない場合は、Word に近い挙動で次 page へ送る。

### 4. Multi-column layout が page 全体に掛かっている

現在の column CSS は page 全体の `article` に付く。そのため、同じ page 内の一部だけ 2カラムにする表現ができない。また、数式・図形・表などの大きい block が CSS multicol の挙動に巻き込まれ、Word と違う分割になりやすい。

必要な挙動:

- columns は page ではなく section layout region に属する。
- まずは equal-width columns を region 単位で正しく扱う。
- `w:col` の unequal-width columns は region model が安定してから追加する。

### 5. `lastRenderedPageBreak` を信用しすぎている

`lastRenderedPageBreak` は Word が保存した時点の page break 記録であって、現在の browser rendering に対する制約ではない。font、image、header/footer、browser metrics、編集後 content が変わると簡単に古くなる。

必要な挙動:

- `lastRenderedPageBreak` は初期分割 hint としてのみ扱う。
- 最終的には実測 overflow で検証する。
- 明示 break と測定結果を、古い saved marker より優先する。

### 6. Overflow split が rendering と密結合している

DOM append 中に overflow を見て、mutable な `breakIndex` を親子へ伝播している。この方式は次のケースで見通しが悪い。

- paragraph / run の途中分割。
- table row continuation。
- header/footer 高さによる content area の変化。
- section break 境界での分割。
- multi-column overflow。

必要な挙動:

- DOM measurement は使い続けるが、split 判断は小さい pagination module に寄せる。
- まず block 境界で安定して split する。
- paragraph / table 内 split は、既存サポート範囲を壊さず段階的に扱う。

### 7. Header/Footer 解決に page context が足りない

header/footer は以下に依存する。

- active section。
- section 内 first page かどうか。
- odd/even page setting。
- `pgNumType` の page number start。
- 前 section からの継承。

現在は物理 page 配列 index と `isFirstPage` に依存しているが、overflow split で page が挿入されると不安定になる。

必要な挙動:

- pagination 後に各 physical page の `PageLayoutContext` を作る。
- context から first/even/default header/footer を解決する。
- 単なる配列 index ではなく、section-local page index を使う。

## 目標モデル

parse 済み document element と DOM rendering の間に、section layout を表す中間モデルを置く。

```ts
interface LayoutRegion {
  section: SectionProperties;
  children: OpenXmlElement[];
  breakBefore?: "none" | "page" | "column" | "evenPage" | "oddPage";
}

interface PhysicalPage {
  regions: LayoutRegion[];
  pageNumber: number;
  sectionPageIndexes: Map<string, number>;
}
```

名前は変えてよいが、重要なのは「1 physical page が複数 section region を持てる」こと。

## 実装計画

### Phase 0: 専用 regression fixture を追加する

先に browser test を追加する。

- `section-continuous-columns`: 1カラム text、continuous section、同じ page 内の 2カラム text。
- `section-next-page`: section break で page size または margin が変わる。
- `section-next-column`: 2カラム section 内の next-column section break。
- `section-header-footer-inheritance`: 複数 section をまたぐ first/even/default header/footer。
- `page-break-stale-last-rendered`: `lastRenderedPageBreak` だけでは正しく分割できない content。

最初の assertion は構造中心にする。

- page count。
- page ごとの region count。
- computed `column-count`。
- page ごとの header/footer text。
- browser page error がないこと。

visual golden は構造が安定してから追加する。

### Phase 1: Section stream を抽出する

document body children を section-scoped region の線形 stream へ変換する小さい module を作る。

責務:

- body order を保持する。
- 各 region に正しい `SectionProperties` を付ける。
- header/footer refs の継承をここで解決する。
- section break type を明示的に保持する。
- この段階では physical page を作らない。

期待結果:

- section boundary が明示的な data になる。
- `splitPageBySymbol` が document 末尾から逆順に `SectionBreak` の意味を補正する必要を減らせる。

### Phase 2: Region-based page rendering を導入する

1 physical page に複数の `article` 相当 region container を render できるようにする。

責務:

- outer page size は physical page context から決める。
- 各 `LayoutRegion` を独立した content container に render する。
- columns は page 全体ではなく region 単位に適用する。
- 単純な 1 section page では既存 DOM とできるだけ同じ出力にする。

この phase では全 edge case を処理しない。重要なのは「同じ page 内に複数 section layout を表現できる」こと。

### Phase 3: 明示 break の意味を整理する

overflow split より前に、明示 break の意味を data として扱う。

- `BreakType.Page`: current physical page を閉じる。
- `BreakType.Column`: current region の column flow で次 column へ進む。
- `SectionType.NextPage`: current page を閉じてから新 section region を始める。
- `SectionType.EvenPage`: 必要なら blank page を挿入してから section を始める。
- `SectionType.OddPage`: 必要なら blank page を挿入してから section を始める。
- `SectionType.Continuous`: current page 上で新 section region を始める。
- `SectionType.NextColumn`: 次 column へ進んでから新 section region を始める。

synthetic DOM の `<s>` element を source of truth にしない。source map / debug marker として残すのはよい。

### Phase 4: Overflow split を region-aware にする

page-only split ではなく、region/page split として扱う。

- region content height を残り page height と比較する。
- region が溢れたら、その region 内で split する。
- 残り content は同じ section properties を持った region として次 page へ送る。
- table continuation と paragraph continuation は既存挙動を壊さない。

`lastRenderedPageBreak` はここで hint として扱う。

- 有効な場合は hinted split を試す。
- render して測定する。
- overflow する、または layout が不正なら measured split に戻す。

### Phase 5: Pagination 後に Header/Footer を解決する

physical page が確定してから実行する。

- physical page number を割り当てる。
- section ごとの page index を計算する。
- first/even/default header/footer refs を解決する。
- 前 section からの refs 継承を適用する。
- header/footer を render / measure して available content height を再計算する。
- content height が変わった page は再 pagination する。

必要なら bounded two-pass にする。

1. header/footer なしで body region を pagination する。
2. header/footer を resolve / measure する。
3. content height が変わった page だけ再 pagination する。
4. 小さい固定回数で止め、収束しない場合は debug 情報を出す。

### Phase 6: Visual golden coverage を追加する

構造テストが通ってから visual または DOM snapshot を追加する。

- column 数が変わる continuous section。
- odd/even header/footer。
- odd/even section break による blank page insertion。
- multi-column region 内の大きな equation / table / drawing。

## 推奨ファイル境界

変更は広げすぎない。

- `src/document/section.ts`: section parse のみ。
- `src/document/page.ts`: page data を region-aware に拡張、または置き換える。
- `src/html-renderer-sync.ts`: DOM rendering は残し、pagination policy は可能な範囲で外へ出す。
- 新規 `src/document/layout.ts`: section stream、layout region、pagination decision。
- tests は `tests/browser` 中心。

図形 rendering や math property 拡張とは混ぜない。別の fidelity track として扱う。

## 完了条件

- 1 physical page が複数 section region を持てる。
- continuous section break で page を強制せず、以降の column 設定が反映される。
- next-page、odd-page、even-page、next-column、page break の明示テストがある。
- overflow で page が増えても first/even/default header/footer が正しく選ばれる。
- `lastRenderedPageBreak` が実測 layout correctness を上書きしない。
- 既存 smoke fixture が browser page error なしで render される。

## リスク

- Browser CSS columns は Word と完全一致しない。region model は差を減らすが、全ては消せない。
- header/footer の測定結果が content capacity を変え、pagination feedback loop を起こす可能性がある。
- table splitting は既に複雑なので、region-aware overflow では既存挙動維持を優先する。
- visual parity は font と browser metrics に依存する。最初は structural tests を優先する。

## 最初の PR 推奨範囲

最初は最小の architecture slice にする。

1. continuous section と next-column section の構造 browser test を追加する。
2. `LayoutRegion` data structure を追加する。
3. document children を section stream に変換する。ただし rendering output はまだ変えない。
4. section stream の unit test を追加する。

最初の PR では overflow split を変えない。review しやすい入力モデルを先に固定し、その後 pagination rewrite に進む。
