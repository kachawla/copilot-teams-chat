import assert from "node:assert/strict";
import test from "node:test";
import { markdownToHtml, textToHtml, toTeamsHtml } from "../com.github.copilot/extensions/teams-chat/lib/markdown.mjs";

test("escapes raw HTML", () => {
    assert.equal(markdownToHtml('<script>alert("x")</script>'), "<p>&lt;script&gt;alert(&quot;x&quot;)&lt;/script&gt;</p>");
});

test("paragraphs and line breaks", () => {
    assert.equal(markdownToHtml("one\ntwo\n\nthree"), "<p>one<br>two</p><p>three</p>");
});

test("inline formatting", () => {
    assert.equal(
        markdownToHtml("**bold** *em* _em2_ ~~gone~~ `a<b>`"),
        "<p><strong>bold</strong> <em>em</em> <em>em2</em> <s>gone</s> <code>a&lt;b&gt;</code></p>",
    );
});

test("does not treat snake_case or math as emphasis", () => {
    assert.equal(markdownToHtml("use my_var_name and 2*3*4"), "<p>use my_var_name and 2*3*4</p>");
});

test("links allow only safe schemes", () => {
    assert.equal(markdownToHtml("[site](https://example.com/a?b=1&c=2)"), '<p><a href="https://example.com/a?b=1&amp;c=2">site</a></p>');
    assert.equal(markdownToHtml("[x](javascript:alert(1))"), "<p>[x](javascript:alert(1))</p>");
    assert.equal(markdownToHtml("[mail](mailto:a@b.com)"), '<p><a href="mailto:a@b.com">mail</a></p>');
});

test("emphasis does not corrupt link hrefs or code", () => {
    assert.equal(
        markdownToHtml("[init](https://x.com/__init__) `__x__`"),
        '<p><a href="https://x.com/__init__">init</a> <code>__x__</code></p>',
    );
});

test("autolinks bare URLs without trailing punctuation", () => {
    assert.equal(
        markdownToHtml("See https://github.com/org/repo/pull/1."),
        '<p>See <a href="https://github.com/org/repo/pull/1">https://github.com/org/repo/pull/1</a>.</p>',
    );
});

test("lists", () => {
    assert.equal(markdownToHtml("- a\n- **b**\n\n1. one\n2. two"), "<ul><li>a</li><li><strong>b</strong></li></ul><ol><li>one</li><li>two</li></ol>");
});

test("fenced code keeps content verbatim and escaped", () => {
    assert.equal(markdownToHtml("```js\nconst a = 1 < 2;\n**not bold**\n```"), "<pre><code>const a = 1 &lt; 2;\n**not bold**</code></pre>");
});

test("headings, quotes, and rules", () => {
    assert.equal(markdownToHtml("# Title\n> quoted *text*\n\n---"), "<h2>Title</h2><blockquote><p>quoted <em>text</em></p></blockquote><hr>");
});

test("placeholder characters in input cannot inject tokens", () => {
    assert.equal(markdownToHtml("a\u00000\u0000b `c`"), "<p>a0b <code>c</code></p>");
});

test("plain text format", () => {
    assert.equal(textToHtml("a **b**\n<c>"), "a **b**<br>&lt;c&gt;");
    assert.equal(toTeamsHtml("**b**", "text"), "**b**");
    assert.equal(toTeamsHtml("**b**"), "<p><strong>b</strong></p>");
});
