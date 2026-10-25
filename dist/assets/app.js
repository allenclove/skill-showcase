/* ============================================================
   Skill 展厅 — 交互
   数据来自 window.SITE_DATA（build.py 从 Excel 生成）
   页面只有两个：首页（技能库）/ 价值
   ============================================================ */
(function () {
  'use strict';

  var D = window.SITE_DATA || {};
  var SITE = D.site || {};
  var SKILLS = (D.skills || []).slice();
  var CASES = (D.cases || []).slice();
  var GUIDE = D.guide || { steps: [], faq: [] };
  var META = D.meta || {};

  var app = document.getElementById('app');
  var modal = document.getElementById('modal');
  var modalBody = document.getElementById('modal-body');
  var toastEl = document.getElementById('toast');

  /* ---------------------------------------------------------- tiny dom */
  function el(tag, props) {
    var node = document.createElement(tag);
    var kids = Array.prototype.slice.call(arguments, 2);
    if (props) {
      Object.keys(props).forEach(function (key) {
        var v = props[key];
        if (v === null || v === undefined || v === false) return;
        if (key === 'class') node.className = v;
        else if (key === 'text') node.textContent = v;
        else if (key === 'dataset') Object.keys(v).forEach(function (d) { node.dataset[d] = v[d]; });
        else if (key.slice(0, 2) === 'on' && typeof v === 'function') node.addEventListener(key.slice(2).toLowerCase(), v);
        else node.setAttribute(key, v === true ? '' : v);
      });
    }
    kids.forEach(function add(kid) {
      if (kid === null || kid === undefined || kid === false) return;
      if (Array.isArray(kid)) return kid.forEach(add);
      node.appendChild(kid.nodeType ? kid : document.createTextNode(String(kid)));
    });
    return node;
  }

  var ICONS = {
    copy: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="11" height="11" rx="2.5"/><path d="M5 15V5.5A2.5 2.5 0 0 1 7.5 3H17"/></svg>',
    check: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>',
    arrow: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M5 12h13M12 5l7 7-7 7"/></svg>',
    search: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.6-3.6"/></svg>',
    spark: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M12 3.5 13.9 9l5.6 1.9-5.6 1.9L12 18.4l-1.9-5.6L4.5 10.9 10.1 9z"/></svg>',
    clock: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round"><circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/></svg>',
    replay: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 1 0 3-6.7"/><path d="M3 4v5h5"/></svg>'
  };

  function icon(name) {
    var span = document.createElement('span');
    span.className = 'ic';
    span.innerHTML = ICONS[name] || '';
    return span;
  }

  function text(str) { return document.createTextNode(String(str)); }

  function fmtMin(m) {
    if (!m && m !== 0) return '—';
    if (m < 60) return m + ' 分钟';
    var h = Math.floor(m / 60), r = m % 60;
    return h + ' 小时' + (r ? ' ' + r + ' 分钟' : '');
  }

  // 省时用「约 X 小时」，避免出现 1 小时 59 分钟这种别扭写法
  function fmtSaved(m) {
    if (!m || m <= 0) return '';
    if (m < 60) return m + ' 分钟';
    return (Math.round((m / 60) * 2) / 2) + ' 小时';
  }

  /* ---------------------------------------------------------- copy */
  async function copyText(value) {
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(value);
        return true;
      }
    } catch (err) { /* fall through */ }
    try {
      var ta = document.createElement('textarea');
      ta.value = value;
      ta.setAttribute('readonly', '');
      ta.style.position = 'fixed';
      ta.style.top = '-1000px';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      ta.setSelectionRange(0, value.length);
      var ok = document.execCommand('copy');
      document.body.removeChild(ta);
      return ok;
    } catch (err2) {
      return false;
    }
  }

  var toastTimer;
  function toast(msg) {
    toastEl.innerHTML = '';
    toastEl.appendChild(icon('check'));
    toastEl.appendChild(text(msg));
    toastEl.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { toastEl.classList.remove('show'); }, 1900);
  }

  function copyButton(value, label, big) {
    return el('button', {
      type: 'button',
      class: 'btn-copy' + (big ? ' btn-big' : ''),
      'data-copy': value,
      'aria-label': '复制提示词'
    }, icon('copy'), el('span', { class: 'copy-label', text: label || '复制提示词' }));
  }

  document.addEventListener('click', async function (e) {
    var btn = e.target.closest('[data-copy]');
    if (!btn) return;
    e.stopPropagation();
    var ok = await copyText(btn.getAttribute('data-copy'));
    if (!ok) { toast('复制失败，请手动选择文本'); return; }

    var msg = btn.getAttribute('data-msg') || '已复制，去对话框粘贴吧';

    // 名称类按钮（标题 / 英文名）：不替换内容，只闪一下
    if (btn.classList.contains('m-copyable')) {
      btn.classList.add('copied');
      toast(msg);
      clearTimeout(btn.__copiedTimer);
      btn.__copiedTimer = setTimeout(function () { btn.classList.remove('copied'); }, 1200);
      return;
    }

    btn.classList.add('ok');
    var label = btn.querySelector('.copy-label');
    var old = label ? label.textContent : '';
    btn.innerHTML = '';
    btn.appendChild(icon('check'));
    btn.appendChild(el('span', { class: 'copy-label', text: '已复制' }));
    toast(msg);
    setTimeout(function () {
      btn.classList.remove('ok');
      btn.innerHTML = '';
      btn.appendChild(icon('copy'));
      btn.appendChild(el('span', { class: 'copy-label', text: old }));
    }, 1800);
  });

  /* ---------------------------------------------------------- 计数动画 */
  function countUp(node, to, dur, digits) {
    var d = digits || 0;
    var start = performance.now();
    var span = Math.max(1, dur || 900);
    function tick(t) {
      var p = Math.min(1, (t - start) / span);
      var e = 1 - Math.pow(1 - p, 3);
      node.textContent = (to * e).toFixed(d);
      if (p < 1) requestAnimationFrame(tick);
      else node.textContent = to.toFixed(d);
    }
    requestAnimationFrame(tick);
  }

  function runCounters(root) {
    Array.prototype.forEach.call(root.querySelectorAll('[data-countup]'), function (node) {
      var to = parseFloat(node.getAttribute('data-countup')) || 0;
      var digits = parseInt(node.getAttribute('data-digits') || '0', 10);
      countUp(node, to, 850, digits);
    });
  }

  /* ---------------------------------------------------------- 对比动画 */
  function playScene(node, force) {
    if (node.dataset.played === '1' && !force) return;
    node.dataset.played = '1';
    node.classList.add('play');

    var beforeNode = node.querySelector('[data-count="before"]');
    var nowNode = node.querySelector('[data-count="now"]');
    var beforeVal = parseFloat(node.dataset.before) || 0;
    var nowVal = parseFloat(node.dataset.now) || 0;

    if (beforeNode) countUp(beforeNode, beforeVal, 1050, 0);
    if (nowNode) {
      nowNode.textContent = '0';
      nowNode.classList.remove('pop');
      setTimeout(function () {
        nowNode.textContent = String(nowVal);
        nowNode.classList.add('pop');
      }, 1150);
    }
  }

  function replayScene(node) {
    if (!node.classList.contains('play')) { playScene(node, true); return; }
    node.classList.remove('play');
    var beforeNode = node.querySelector('[data-count="before"]');
    var nowNode = node.querySelector('[data-count="now"]');
    if (beforeNode) beforeNode.textContent = '0';
    if (nowNode) { nowNode.classList.remove('pop'); nowNode.textContent = '0'; }
    void node.offsetWidth;
    setTimeout(function () { playScene(node, true); }, 40);
  }

  function initScenes(root) {
    var nodes = Array.prototype.slice.call(root.querySelectorAll('.scene'));
    if (!nodes.length) return;

    if (!('IntersectionObserver' in window)) {
      nodes.forEach(function (n) { playScene(n, true); });
      return;
    }

    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (entry) {
        if (!entry.isIntersecting) return;
        var node = entry.target;
        io.unobserve(node);
        playScene(node);
      });
    }, { threshold: 0.18, rootMargin: '0px 0px -6% 0px' });

    nodes.forEach(function (n) { io.observe(n); });
  }

  /* ---------------------------------------------------------- shared */
  function emphasis(title) {
    var m = String(title).match(/^([\s\S]*?)「([^」]+)」([\s\S]*)$/);
    if (!m) return [text(title)];
    return [text(m[1]), el('em', null, m[2]), text(m[3])];
  }

  function statusTag(status) {
    var map = { '已上线': 'ok', '试点': 'warn', '规划': 'plan' };
    return el('span', { class: 'status ' + (map[status] || 'plan') }, el('i'), text(status));
  }

  function sectionHead(eyebrow, title, desc, moreHref, moreLabel) {
    var left = el('div', null,
      el('div', { class: 'eyebrow', text: eyebrow }),
      el('h2', { class: 'sec-title', text: title }),
      desc ? el('p', { class: 'sec-desc', text: desc }) : null
    );
    var right = moreHref ? el('a', { class: 'more', href: moreHref }, text(moreLabel || '查看全部'), icon('arrow')) : null;
    return el('div', { class: 'shead' }, left, right);
  }

  function pageHead(title, desc) {
    return el('div', { class: 'wrap' },
      el('header', { class: 'phead' },
        el('h1', { text: title }),
        desc ? el('p', { text: desc }) : null
      )
    );
  }

  /* ---------------------------------------------------------- 技能卡 */
  function skillCard(s) {
    var card = el('article', {
      class: 'card',
      tabindex: '0',
      role: 'button',
      'aria-label': s.name + '，查看详情',
      onclick: function (e) {
        if (e.target.closest('[data-copy]')) return;
        openSkill(s);
      },
      onkeydown: function (e) {
        if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); openSkill(s); }
      }
    });

    card.appendChild(el('div', { class: 'card-top' },
      el('span', { class: 'chip', text: s.biz }),
      el('span', { class: 'chip chip-line', text: s.func }),
      statusTag(s.status)
    ));

    card.appendChild(el('h3', { class: 'card-title', text: s.name }));
    if (s.en) card.appendChild(el('p', { class: 'card-en', text: s.en }));
    card.appendChild(el('p', { class: 'card-desc', text: s.desc }));

    if (s.gain) card.appendChild(el('span', { class: 'card-gain' }, icon('clock'), text(s.gain)));

    return card;
  }

  /* ---------------------------------------------------------- 价值场景 */
  function sceneCard(c) {
    var beforeVal = +c.before || 0;
    var nowVal = +c.now || 0;
    var wnow = Math.max(1.2, (nowVal / Math.max(beforeVal, 1)) * 100);
    var pct = beforeVal > 0 ? Math.round(((beforeVal - nowVal) / beforeVal) * 100) : 0;

    var node = el('article', {
      class: 'scene',
      style: '--wnow:' + wnow.toFixed(2) + '%'
    });
    node.dataset.before = beforeVal;
    node.dataset.now = nowVal;

    var replay = el('button', {
      type: 'button',
      class: 'scene-replay',
      'aria-label': '重播这段对比动画',
      onclick: function () { replayScene(node); }
    }, icon('replay'), text('重播'));

    node.appendChild(el('div', { class: 'scene-hd' },
      el('span', { class: 'scene-no', text: c.id || '' }),
      el('h3', { text: c.title }),
      c.role ? el('span', { class: 'chip chip-plain role', text: c.role }) : null,
      replay
    ));

    var steps = el('ol', { class: 'steps' }, (c.beforeSteps || []).map(function (st, i) {
      return el('li', { style: '--i:' + i, text: st });
    }));

    var beforePane = el('div', { class: 'pane pane-before' },
      el('div', { class: 'pane-hd' },
        el('span', { class: 'dot' }),
        text('以前'),
        el('span', { class: 'clock' },
          el('b', { 'data-count': 'before', text: '0' }),
          el('small', { text: '分钟' })
        )
      ),
      steps
    );

    var nowPane = el('div', { class: 'pane pane-now' },
      el('div', { class: 'pane-hd' },
        el('span', { class: 'dot' }),
        text('现在'),
        el('span', { class: 'clock ok' },
          el('b', { 'data-count': 'now', text: '0' }),
          el('small', { text: '分钟' })
        )
      ),
      el('div', { class: 'type-wrap' }, el('pre', { class: 'prompt', text: c.nowPrompt })),
      el('div', { class: 'pane-actions' },
        copyButton(c.nowPrompt, '复制这一句', true),
        el('span', { class: 'now-badge' }, icon('check'), text('一句话搞定'))
      )
    );

    node.appendChild(el('div', { class: 'scene-body' },
      beforePane,
      el('div', { class: 'scene-arrow' }, icon('arrow')),
      nowPane
    ));

    node.appendChild(el('div', { class: 'scene-ft' },
      el('div', { class: 'bars' },
        el('div', { class: 'bar-row' },
          el('span', { class: 'lb', text: '以前' }),
          el('div', { class: 'bar-track bar-before' }, el('i')),
          el('span', { class: 'v', text: beforeVal + ' min' })
        ),
        el('div', { class: 'bar-row' },
          el('span', { class: 'lb', text: '现在' }),
          el('div', { class: 'bar-track bar-now' }, el('i')),
          el('span', { class: 'v', text: nowVal + ' min' })
        )
      ),
      el('span', { class: 'save' },
        el('b', { text: '- ' + pct + '%' }),
        text('省约 ' + fmtSaved(beforeVal - nowVal))
      ),
      el('p', { class: 'scene-value' }, icon('spark'), text(c.value))
    ));

    node.classList.add('anim');
    return node;
  }

  /* ---------------------------------------------------------- 首页 */
  function heroBlock() {
    var saved = META.savedMinutes || 0;
    var items = [
      { v: SKILLS.length, d: 0, u: '个', l: '技能' },
      { v: META.bizCount || 0, d: 0, u: '类', l: '业务' },
      { v: META.funcCount || 0, d: 0, u: '类', l: '功能' },
      { v: saved / 60, d: 1, u: 'h', l: '单次合计省时' }
    ];

    var title = el('h1', { class: 'hero-title' });
    emphasis(SITE.title).forEach(function (n) { title.appendChild(n); });

    return el('section', { class: 'hero' },
      el('div', { class: 'wrap' },
        el('div', { class: 'hero-top' },
          el('div', { class: 'hero-copy' },
            el('span', { class: 'hero-badge' }, el('span', { class: 'pip' }), text(SITE.badge)),
            title,
            SITE.desc ? el('p', { class: 'hero-desc', text: SITE.desc }) : null
          ),
          el('div', { class: 'hero-stats' }, items.map(function (s) {
            return el('div', { class: 'hs' },
              el('b', null,
                el('span', { 'data-countup': s.v, 'data-digits': s.d, text: '0' }),
                el('small', { text: s.u })
              ),
              el('span', { class: 'hs-l', text: s.l })
            );
          }))
        )
      )
    );
  }

  function libraryBlock() {
    var state = { q: '', biz: '全部', func: '全部' };

    var searchInput = el('input', {
      type: 'search',
      placeholder: '搜技能、场景、负责人或标签…',
      'aria-label': '搜索技能',
      oninput: function (e) { state.q = e.target.value; apply(); }
    });

    var searchBox = el('div', { class: 'search' }, icon('search'), searchInput, el('kbd', { text: '/' }));

    var resultLine = el('div', { class: 'result-line' });

    var sticky = el('div', { class: 'lib-search' }, searchBox, resultLine);

    var grid = el('div', { class: 'grid' });

    var bizRow = chipRow('业务', 'biz', META.bizCategories || [], state, apply);
    var funcRow = chipRow('功能', 'func', META.funcCategories || [], state, apply);

    function apply() {
      var q = state.q.trim().toLowerCase();
      var list = SKILLS.filter(function (s) {
        if (state.biz !== '全部' && s.biz !== state.biz) return false;
        if (state.func !== '全部' && s.func !== state.func) return false;
        if (!q) return true;
        var hay = [s.name, s.desc, s.biz, s.func, s.ownerBiz, s.ownerIT, s.value,
          (s.tags || []).join(' '), (s.prompts || []).join(' '), s.scenario].join(' ').toLowerCase();
        return hay.indexOf(q) !== -1;
      });

      grid.innerHTML = '';
      resultLine.innerHTML = '';
      resultLine.appendChild(el('strong', { text: list.length }));
      resultLine.appendChild(text(' / ' + SKILLS.length + ' 个'));

      if (!list.length) {
        grid.appendChild(el('div', { class: 'empty' },
          el('h3', { text: '没有匹配的技能' }),
          el('p', { text: '换个关键词，或清掉筛选条件再看一次。' }),
          el('div', { style: 'margin-top:16px' },
            el('button', {
              class: 'btn btn-quiet', type: 'button',
              onclick: function () {
                state.q = ''; state.biz = '全部'; state.func = '全部';
                searchInput.value = '';
                bizRow.reset(); funcRow.reset();
                apply();
              }
            }, text('清除筛选'))
          )
        ));
        return;
      }

      list.forEach(function (s) { grid.appendChild(skillCard(s)); });
    }

    var section = el('section', { class: 'lib', id: 'lib' },
      el('div', { class: 'wrap' },
        el('div', { class: 'lib-hd' },
          el('h2', null, text('技能库'), el('span', { class: 'count', text: SKILLS.length + ' 个技能 · 点卡片看详情' }))
        ),
        sticky,
        bizRow, funcRow,
        grid
      )
    );

    setTimeout(apply, 0);
    return section;
  }

  function chipRow(label, field, values, state, apply) {
    var row = el('div', { class: 'filter-row' }, el('span', { class: 'flabel', text: label }));
    var buttons = {};

    var all = el('button', { type: 'button', class: 'fchip on' },
      text('全部'), el('b', { text: SKILLS.length }));
    buttons['全部'] = all;
    row.appendChild(all);

    values.forEach(function (v) {
      var count = SKILLS.filter(function (s) { return s[field] === v; }).length;
      var b = el('button', { type: 'button', class: 'fchip' }, text(v), el('b', { text: count }));
      buttons[v] = b;
      row.appendChild(b);
    });

    Object.keys(buttons).forEach(function (key) {
      buttons[key].addEventListener('click', function () {
        state[field] = key;
        Object.keys(buttons).forEach(function (k) { buttons[k].classList.toggle('on', k === key); });
        apply();
      });
    });

    row.reset = function () {
      state[field] = '全部';
      Object.keys(buttons).forEach(function (k) { buttons[k].classList.toggle('on', k === '全部'); });
    };

    return row;
  }

  function bandBlock() {
    return el('section', { class: 'sec sec-tight' },
      el('div', { class: 'wrap' },
        el('div', { class: 'band' },
          el('h2', null, text('把这份网页发给同事，'), el('br'), text('比在会上讲十遍管用。')),
          el('p', { text: '整个站点是本地文件，双击就能看：不用登录、不用联网、不用装环境。' }),
          el('div', { class: 'hero-actions' },
            el('a', { class: 'btn btn-primary', href: '#/story' }, text('看价值故事'), icon('arrow')),
            el('a', { class: 'btn btn-ghost', href: '#/guide' }, text('怎么用'))
          )
        )
      )
    );
  }

  function pageHome() {
    var page = el('div', { class: 'page' });
    page.appendChild(heroBlock());
    page.appendChild(libraryBlock());
    return page;
  }

  /* ---------------------------------------------------------- 价值页 */
  function pageStory() {
    var saved = META.savedMinutes || 0;
    var avg = 0;
    if (CASES.length) {
      avg = Math.round(CASES.reduce(function (acc, c) {
        return acc + (c.before > 0 ? ((c.before - c.now) / c.before) * 100 : 0);
      }, 0) / CASES.length);
    }

    var page = el('div', { class: 'page' });

    page.appendChild(el('section', { class: 'story-hero' },
      el('div', { class: 'wrap' },
        el('span', { class: 'hero-badge' }, el('span', { class: 'pip' }), text('BEFORE / AFTER · 逐个对照')),
        el('h1', { class: 'sh-title' },
          text('同一件活，'),
          el('span', { class: 'sh-old', text: '以前忙半天' }),
          text('，'),
          el('br'),
          text('现在'),
          el('span', { class: 'sh-new', text: '两句话' }),
          text('。')
        ),
        el('p', { class: 'sh-desc', text: '下面每一段都是拿真实任务掐表记的：左边把以前的步骤一条条摆出来，右边是现在真正发出去的那句话。' }),
        el('div', { class: 'sh-stats' },
          el('div', null,
            el('div', { class: 'sh-num' }, el('span', { 'data-countup': CASES.length, 'data-digits': 0, text: '0' }), el('small', { text: '个' })),
            el('div', { class: 'sh-lb', text: '真实场景' })
          ),
          el('div', null,
            el('div', { class: 'sh-num' }, el('span', { 'data-countup': avg, 'data-digits': 0, text: '0' }), el('small', { text: '%' })),
            el('div', { class: 'sh-lb', text: '平均时间降幅' })
          ),
          el('div', null,
            el('div', { class: 'sh-num' }, el('span', { 'data-countup': saved / 60, 'data-digits': 1, text: '0' }), el('small', { text: 'h' })),
            el('div', { class: 'sh-lb', text: '单次合计省时' })
          )
        )
      )
    ));

    page.appendChild(el('section', { class: 'sec-tight' },
      el('div', { class: 'wrap' },
        el('div', { class: 'story-list' }, CASES.map(sceneCard)),
        el('p', { class: 'sec-desc', style: 'margin-top:22px', text: '注：耗时按实际任务记录，含找模板、返工和确认环节；不同熟练度会有差异，但量级不会变。' })
      )
    ));

    page.appendChild(bandBlock());
    return page;
  }

  /* ---------------------------------------------------------- 指南 */
  function pageGuide() {
    var page = el('div', { class: 'page' });
    page.appendChild(pageHead('使用指南', '不用培训、不用装软件。会打字，就会用。'));

    page.appendChild(el('section', { class: 'sec-tight' },
      el('div', { class: 'wrap' },
        el('div', { class: 'gsteps' }, (GUIDE.steps || []).map(function (s, i) {
          return el('div', { class: 'gstep' },
            el('span', { class: 'n', text: '0' + (i + 1) }),
            el('h3', { text: s.title }),
            el('p', { text: s.body })
          );
        }))
      ),
      el('div', { class: 'wrap', style: 'margin-top:40px' },
        sectionHead('FAQ', '常见问题', null, null, null),
        el('div', { class: 'faq' }, (GUIDE.faq || []).map(function (f) {
          return el('details', null,
            el('summary', { text: f.title }),
            el('div', { class: 'ans', text: f.body })
          );
        }))
      )
    ));

    page.appendChild(el('section', { class: 'sec sec-alt' },
      el('div', { class: 'wrap' },
        sectionHead('MAINTAIN', '给维护者的三步', '这份网页是「数据 + 模板」生成的，不需要写代码。'),
        el('div', { class: 'gsteps' },
          el('div', { class: 'gstep' }, el('span', { class: 'n', text: '01' }),
            el('h3', { text: '改 Excel' }),
            el('p', { text: '打开《技能展厅.xlsx》，在「技能清单」里增删改：业务分类、功能分类、业务负责人、IT 负责人、提示词示例。' })),
          el('div', { class: 'gstep' }, el('span', { class: 'n', text: '02' }),
            el('h3', { text: '双击生成' }),
            el('p', { text: '双击「一键生成.cmd」。脚本会检查必填项和编号重复，有问题会指到具体行，没问题就重建 dist 目录。' })),
          el('div', { class: 'gstep' }, el('span', { class: 'n', text: '03' }),
            el('h3', { text: '发出去' }),
            el('p', { text: '把整个 dist 文件夹打包发给同事，解压后双击 index.html 即可。对外分享前记得清掉内部信息。' }))
        )
      )
    ));

    return page;
  }

  /* ---------------------------------------------------------- 详情抽屉 */
  function openSkill(skill) {
    modalBody.innerHTML = '';

    var body = el('div', { class: 'm-body' });

    body.appendChild(el('div', { class: 'm-kicker' }, statusTag(skill.status)));

    body.appendChild(el('h2', { class: 'm-title', id: 'modal-title' },
      el('button', {
        type: 'button', class: 'm-copyable',
        'data-copy': skill.name, 'data-msg': '已复制技能名称',
        title: '点击复制技能名称',
        'aria-label': '复制技能名称 ' + skill.name
      }, text(skill.name), icon('copy'))
    ));

    if (skill.en) {
      body.appendChild(el('p', { class: 'm-en' },
        el('button', {
          type: 'button', class: 'm-copyable',
          'data-copy': skill.en, 'data-msg': '已复制英文名称',
          title: '点击复制英文名称',
          'aria-label': '复制英文名称 ' + skill.en
        }, text(skill.en), icon('copy'))
      ));
    }

    body.appendChild(el('p', { class: 'm-desc', text: skill.desc }));

    /* 重点 1：提示词 */
    var promptSec = el('section', { class: 'm-sec m-hero' },
      el('div', { class: 'm-sec-hd' },
        el('span', { class: 'm-label', text: '提示词示例' }),
        el('span', { class: 'm-hint', text: '复制后把括号里的信息换成你自己的' })
      )
    );

    (skill.prompts || []).forEach(function (p, i) {
      promptSec.appendChild(el('div', { class: 'm-prompt' },
        el('div', { class: 'm-prompt-hd' },
          el('span', { text: skill.prompts.length > 1 ? '示例 ' + (i + 1) : '可直接复制' }),
          copyButton(p, '复制')
        ),
        el('pre', { class: 'm-prompt-text', text: p })
      ));
    });
    body.appendChild(promptSec);

    /* 重点 2：价值 */
    body.appendChild(el('section', { class: 'm-sec' },
      el('div', { class: 'm-sec-hd' }, el('span', { class: 'm-label', text: '提升价值' })),
      el('div', { class: 'm-value' },
        skill.gain ? el('span', { class: 'm-gain' }, icon('clock'), text(skill.gain)) : null,
        el('p', { class: 'm-value-text', text: skill.value })
      )
    ));

    /* 次要信息 */
    var soft = el('section', { class: 'm-sec m-soft' });
    if (skill.scenario) {
      soft.appendChild(el('div', { class: 'm-meta' },
        el('h4', { text: '适用场景' }),
        el('p', { text: skill.scenario })
      ));
    }
    var taglines = el('div', { class: 'm-taglines' });
    if (skill.biz) {
      taglines.appendChild(el('div', { class: 'm-tagrow' },
        el('h4', { text: '业务标签' }),
        el('div', { class: 'm-tags' }, el('span', { class: 'tag tag-biz', text: skill.biz }))
      ));
    }
    if (skill.func) {
      taglines.appendChild(el('div', { class: 'm-tagrow' },
        el('h4', { text: '功能标签' }),
        el('div', { class: 'm-tags' }, el('span', { class: 'tag', text: skill.func }))
      ));
    }
    soft.appendChild(taglines);

    soft.appendChild(el('div', { class: 'm-owners' },
      el('span', null, el('em', { text: '业务负责人' }), text(skill.ownerBiz)),
      el('span', null, el('em', { text: 'IT 负责人' }), text(skill.ownerIT))
    ));

    body.appendChild(soft);

    modalBody.appendChild(body);

    modal.classList.add('open');
    modal.setAttribute('aria-hidden', 'false');
    document.body.classList.add('no-scroll');
    modalBody.scrollTop = 0;
    var card = modal.querySelector('.modal-card');
    if (card) card.focus({ preventScroll: true });
  }

  function closeModal() {
    if (!modal.classList.contains('open')) return;
    modal.classList.remove('open');
    modal.setAttribute('aria-hidden', 'true');
    document.body.classList.remove('no-scroll');
  }

  modal.addEventListener('click', function (e) {
    if (e.target.closest('[data-close]')) closeModal();
  });

  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') closeModal();
    if (e.key !== '/') return;
    var active = document.activeElement;
    if (active && (/^(INPUT|TEXTAREA|SELECT)$/.test(active.tagName) || active.isContentEditable)) return;
    var input = document.querySelector('.search input');
    if (!input) return;
    e.preventDefault();
    input.focus();
  });

  /* ---------------------------------------------------------- 路由 */
  var ROUTES = {
    home: { title: '', render: pageHome },
    story: { title: '价值', render: pageStory },
    guide: { title: '使用指南', render: pageGuide }
  };

  function parseHash() {
    var raw = (location.hash || '').replace(/^#\/?/, '');
    var name = (raw.split('/')[0] || 'home').toLowerCase();
    if (name === 'skills') name = 'home';
    return ROUTES[name] ? name : 'home';
  }

  function render() {
    var name = parseHash();
    var route = ROUTES[name];

    closeModal();
    app.innerHTML = '';
    app.classList.remove('page-enter');
    void app.offsetWidth;
    app.classList.add('page-enter');
    app.appendChild(route.render());

    document.title = (route.title ? route.title + ' · ' : '') + SITE.name;

    Array.prototype.forEach.call(document.querySelectorAll('[data-nav]'), function (a) {
      var on = a.dataset.nav === name;
      a.classList.toggle('active', on);
      if (on) a.setAttribute('aria-current', 'page');
      else a.removeAttribute('aria-current');
    });

    runCounters(app);
    initScenes(app);
    window.scrollTo(0, 0);
  }

  window.addEventListener('hashchange', render);

  var contact = document.getElementById('ftr-contact');
  if (contact && SITE.contact) contact.textContent = ' · ' + SITE.contact;

  render();
})();
