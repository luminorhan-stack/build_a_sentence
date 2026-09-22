/* TOEFL Writing Build a Sentence 练习系统 主逻辑
 * 依赖：build-a-sentence-data.js（须先于本文件加载，提供全局 BANK_VERSION / EMBEDDED_BANK）
 *       xlsx.full.min.js（CDN，题目管理页 Excel 导入用）
 * 逻辑与 v3 内联版逐行一致，未做任何改写
 */

        /* ================================================================
         * 全局状态
         * ================================================================ */
        let questions = [];            // 当前题库（内置或自定义导入）
        let usingCustomBank = false;
        let currentQuestionIndex = 0;
        let currentQuestion = null;
        let slots = [];                // 槽位状态（长度 = 模板槽位数）
        let usedWords = new Set();
        let stats = {
            total: 0,
            correct: 0,
            streak: 0,
            history: []
        };

        // 练习池：筛选 / 套题 / 错题重练 共用同一练习界面
        let pool = { list: [], mode: 'bank', label: '' };

        // 套题进行中的会话
        let setSession = null;         // { setId, list, answers: {qid: bool} }
        let setStats = {};             // { setId: { attempts: [{at,correct,total}], inProgress: {qid:bool} } }

        // 错题本
        const WB_KEY = 'toeflBuildWrongbook';
        const SETS_KEY = 'toeflBuildSets';
        const WB_REASONS = ['语序错误', '句子结构', '固定搭配', '语境理解偏差', '粗心失误', '其他'];
        const WB_STATUSES = ['未掌握', '复习中', '已掌握'];
        let wrongBookItems = [];
        let wbEditingId = null;
        let wbDeleteArmed = null;
        let wbDeleteTimer = null;

        // DOM Elements
        const uploadSection = document.getElementById('uploadSection');
        const fileInput = document.getElementById('fileInput');
        const questionContainer = document.getElementById('questionContainer');
        const emptyState = document.getElementById('emptyState');
        const promptText = document.getElementById('promptText');
        const sentenceContainer = document.getElementById('sentenceContainer');
        const optionsArea = document.getElementById('optionsArea');
        const currentNum = document.getElementById('currentNum');
        const totalNum = document.getElementById('totalNum');
        const progressFill = document.getElementById('progressFill');
        const feedback = document.getElementById('feedback');

        /* ================================================================
         * 初始化
         * ================================================================ */
        document.addEventListener('DOMContentLoaded', () => {
            loadStats();
            loadWrongBook();
            loadSetStats();
            initBank();
            setupEventListeners();
            buildFilterOptions();
            resetToBankPool();
        });

        function initBank() {
            const flag = localStorage.getItem('toeflBuildBankFlag');
            if (flag === 'custom') {
                try {
                    const saved = JSON.parse(localStorage.getItem('toeflBuildQuestions') || 'null');
                    if (saved && Array.isArray(saved) && saved.length > 0) {
                        questions = saved;
                        usingCustomBank = true;
                        return;
                    }
                } catch (e) { /* fallthrough */ }
            }
            questions = EMBEDDED_BANK.map(q => Object.assign({}, q));
            usingCustomBank = false;
        }

        function persistBank() {
            // 仅自定义题库写入 localStorage；内置题库不占配额
            try {
                localStorage.setItem('toeflBuildQuestions', JSON.stringify(questions));
                localStorage.setItem('toeflBuildBankFlag', 'custom');
                usingCustomBank = true;
            } catch (e) {
                showToast('保存题库失败：本地存储空间不足');
            }
        }

        /* ================================================================
         * 事件绑定
         * ================================================================ */
        function setupEventListeners() {
            // Tab navigation
            document.querySelectorAll('.nav-btn').forEach(btn => {
                btn.addEventListener('click', () => {
                    document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
                    document.querySelectorAll('.tab-content').forEach(t => t.classList.remove('active'));
                    btn.classList.add('active');
                    document.getElementById(`${btn.dataset.tab}-tab`).classList.add('active');
                    if (btn.dataset.tab === 'stats') updateStatsDisplay();
                    if (btn.dataset.tab === 'manage') renderQuestionList();
                    if (btn.dataset.tab === 'sets') renderSets();
                    if (btn.dataset.tab === 'wrongbook') renderWrongBook();
                    if (btn.dataset.tab === 'banktable') renderBankTable();
                });
            });

            // File upload
            uploadSection.addEventListener('click', () => fileInput.click());
            uploadSection.addEventListener('dragover', handleDragOver);
            uploadSection.addEventListener('dragleave', handleDragLeave);
            uploadSection.addEventListener('drop', handleDrop);
            fileInput.addEventListener('change', handleFileSelect);

            // Practice buttons
            document.getElementById('resetBtn').addEventListener('click', resetQuestion);
            document.getElementById('prevBtn').addEventListener('click', prevQuestion);
            document.getElementById('nextBtn').addEventListener('click', nextQuestion);
            document.getElementById('checkBtn').addEventListener('click', checkAnswer);
            document.getElementById('reuploadBtn').addEventListener('click', reuploadFile);
            document.getElementById('bannerExit').addEventListener('click', exitSpecialMode);

            // 筛选
            ['fScenario', 'fPolarity', 'fFunction', 'fStructure', 'fPurpose'].forEach(id => {
                document.getElementById(id).addEventListener('change', applyFilters);
            });
            document.getElementById('fKeyword').addEventListener('input', debounce(applyFilters, 250));
            document.getElementById('fReset').addEventListener('click', resetAllFilters);

            // Manage
            document.getElementById('searchBox').addEventListener('input', renderQuestionList);
            document.getElementById('addQuestionBtn').addEventListener('click', openAddModal);
            document.getElementById('exportQuestionsBtn').addEventListener('click', exportQuestions);
            document.getElementById('clearAllQuestionsBtn').addEventListener('click', restoreBank);

            // Stats
            document.getElementById('clearStatsBtn').addEventListener('click', clearStats);
            document.getElementById('exportStatsBtn').addEventListener('click', exportStats);

            // Modal
            document.getElementById('modalClose').addEventListener('click', closeModal);
            document.getElementById('modalCancelBtn').addEventListener('click', closeModal);
            document.getElementById('questionForm').addEventListener('submit', saveQuestion);

            // 错题本
            ['wbScenarioFilter', 'wbFunctionFilter', 'wbPurposeFilter', 'wbDateFilter', 'wbStatusFilter'].forEach(id => {
                document.getElementById(id).addEventListener('change', renderWbList);
            });
            document.getElementById('wbSearch').addEventListener('input', debounce(renderWbList, 250));
            document.getElementById('wbPracticeBtn').addEventListener('click', startWrongbookPractice);
            document.getElementById('wbExportBtn').addEventListener('click', exportWrongBook);
            document.getElementById('wbImportBtn').addEventListener('click', () => document.getElementById('wbImportFile').click());
            document.getElementById('wbImportFile').addEventListener('change', importWrongBook);
            document.getElementById('wbEditClose').addEventListener('click', closeWbEdit);
            document.getElementById('wbEditCancel').addEventListener('click', closeWbEdit);
            document.getElementById('wbEditSave').addEventListener('click', saveWbEdit);

            // 套题成绩弹窗
            document.getElementById('summaryClose').addEventListener('click', closeSetSummary);
            document.getElementById('summaryRestart').addEventListener('click', () => {
                const setId = document.getElementById('summaryRestart').dataset.setId;
                closeSetSummary();
                startSet(setId, true);
            });
            document.getElementById('summaryWrong').addEventListener('click', () => {
                const wrongIds = JSON.parse(document.getElementById('summaryWrong').dataset.wrong || '[]');
                closeSetSummary();
                startReviewPractice(wrongIds, '套题错题重练');
            });
            document.getElementById('summaryBack').addEventListener('click', () => {
                closeSetSummary();
                exitSpecialMode();
                switchTab('sets');
            });
        }

        function switchTab(tab) {
            document.querySelector(`.nav-btn[data-tab="${tab}"]`).click();
        }

        /* ================================================================
         * 题库总表（数据库视图）
         * ================================================================ */
        const BT_PAGE_SIZE = 100;
        let btState = { page: 1, sortKey: 'id', sortDir: 1, filters: { kw: '', source: '', scenario: '', polarity: '', func: '', structure: '', purpose: '' } };

        function btInit() {
            const dims = [
                ['btSource', 'source', '套题'], ['btScenario', 'scenario', '场景'],
                ['btPolarity', 'polarity', '肯定/否定'], ['btFunction', 'func', '句子功能'],
                ['btStructure', 'structure', '句子结构'], ['btPurpose', 'purpose', '交际目的']
            ];
            dims.forEach(([selId, key, label]) => {
                const el = document.getElementById(selId);
                const vals = [...new Set(questions.map(q => (q[keyMap(key)] || '').trim()).filter(Boolean))].sort();
                el.innerHTML = `<option value="">全部${label}</option>` + vals.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
                el.addEventListener('change', () => { btState.filters[key] = el.value; btState.page = 1; renderBankTable(); });
            });
            const kw = document.getElementById('btKeyword');
            kw.addEventListener('input', debounce(() => { btState.filters.kw = kw.value.trim().toLowerCase(); btState.page = 1; renderBankTable(); }, 250));
            document.getElementById('btReset').addEventListener('click', () => {
                btState = { page: 1, sortKey: 'id', sortDir: 1, filters: { kw: '', source: '', scenario: '', polarity: '', func: '', structure: '', purpose: '' } };
                kw.value = '';
                dims.forEach(([selId]) => { document.getElementById(selId).value = ''; });
                renderBankTable();
            });
        }

        function keyMap(k) {
            return { source: 'source', scenario: 'scenario', polarity: 'polarity', func: 'sentenceFunction', structure: 'sentenceStructure', purpose: 'purpose' }[k];
        }

        function btFiltered() {
            const f = btState.filters;
            return questions.filter(q => {
                if (f.source && q.source !== f.source) return false;
                if (f.scenario && q.scenario !== f.scenario) return false;
                if (f.polarity && q.polarity !== f.polarity) return false;
                if (f.func && q.sentenceFunction !== f.func) return false;
                if (f.structure && q.sentenceStructure !== f.structure) return false;
                if (f.purpose && q.purpose !== f.purpose) return false;
                if (f.kw) {
                    const hay = `${q.id} ${q.prompt || ''} ${q.questionText || ''} ${q.correctAnswer || ''} ${q.explanation || ''}`.toLowerCase();
                    if (!hay.includes(f.kw)) return false;
                }
                return true;
            });
        }

        function btSortVal(q, key) {
            switch (key) {
                case 'id': return q.id;
                case 'date': return q.date || '';
                case 'source': return q.source || '';
                case 'scenario': return q.scenario || '';
                case 'polarity': return q.polarity || '';
                case 'func': return q.sentenceFunction || '';
                case 'structure': return q.sentenceStructure || '';
                case 'grammar': return q.grammar || '';
                case 'purpose': return q.purpose || '';
                case 'sample': return q.correctAnswer || '';
                default: return '';
            }
        }

        function renderBankTable() {
            if (!btState.inited) { btInit(); btState.inited = true; }
            const wrap = document.getElementById('btTableWrap');
            const list = btFiltered();
            const key = btState.sortKey, dir = btState.sortDir;
            list.sort((a, b) => btSortVal(a, key).localeCompare(btSortVal(b, key), 'zh-CN') * dir);

            const totalPages = Math.max(1, Math.ceil(list.length / BT_PAGE_SIZE));
            if (btState.page > totalPages) btState.page = totalPages;
            const start = (btState.page - 1) * BT_PAGE_SIZE;
            const pageItems = list.slice(start, start + BT_PAGE_SIZE);

            document.getElementById('btCount').textContent = `${list.length} 题 · 第 ${btState.page}/${totalPages} 页`;

            const cols = [
                ['id', '题号'], ['date', '日期'], ['source', '套题'], ['scenario', '场景'],
                ['polarity', '肯/否'], ['func', '句子功能'], ['structure', '句子结构'],
                ['grammar', '语法'], ['purpose', '交际目的'], ['sample', '范答']
            ];
            const chip = (v, cls) => v ? `<span class="bt-chip ${cls || ''}">${escapeHtml(v)}</span>` : '<span class="bt-empty">—</span>';
            let html = '<table class="bt-table"><thead><tr>' + cols.map(([k, label]) => {
                const ind = btState.sortKey === k ? (btState.sortDir > 0 ? '<span class="sort-ind">▲</span>' : '<span class="sort-ind">▼</span>') : '';
                return `<th data-key="${k}">${label}${ind}</th>`;
            }).join('') + '</tr></thead><tbody>';
            pageItems.forEach((q, i) => {
                html += `<tr class="${i % 2 ? '' : 'bt-row-odd'}" onclick="goToQuestion('${escapeJs(q.id)}')" title="点击进入练习">`
                    + `<td class="bt-id">${escapeHtml(q.id)}</td>`
                    + `<td>${escapeHtml(q.date || '—')}</td>`
                    + `<td>${escapeHtml(q.source || '—')}</td>`
                    + `<td>${chip(q.scenario)}</td>`
                    + `<td>${chip(q.polarity, 'c-blue')}</td>`
                    + `<td>${chip(q.sentenceFunction)}</td>`
                    + `<td>${chip(q.sentenceStructure, 'c-brown')}</td>`
                    + `<td>${chip(q.grammar, 'c-gold')}</td>`
                    + `<td>${chip(q.purpose)}</td>`
                    + `<td class="bt-sample" title="${escapeHtml(q.correctAnswer || '')}">${escapeHtml(q.correctAnswer || '—')}</td>`
                    + '</tr>';
            });
            html += '</tbody></table>';
            if (!list.length) html = '<div class="empty-state"><div class="empty-icon">🗂️</div><h3>无匹配题目</h3><p>调整筛选条件后重试</p></div>';
            wrap.innerHTML = html;

            wrap.querySelectorAll('th[data-key]').forEach(th => {
                th.addEventListener('click', () => {
                    const k = th.dataset.key;
                    if (btState.sortKey === k) btState.sortDir *= -1; else { btState.sortKey = k; btState.sortDir = 1; }
                    renderBankTable();
                });
            });

            const pager = document.getElementById('btPager');
            let ph = `<button ${btState.page === 1 ? 'disabled' : ''} data-pg="${btState.page - 1}">‹ 上一页</button>`;
            const maxBtn = 9;
            let p1 = Math.max(1, btState.page - 4), p2 = Math.min(totalPages, p1 + maxBtn - 1);
            p1 = Math.max(1, p2 - maxBtn + 1);
            if (p1 > 1) ph += `<button data-pg="1">1</button><span class="bt-pager-info">…</span>`;
            for (let p = p1; p <= p2; p++) ph += `<button class="${p === btState.page ? 'cur' : ''}" data-pg="${p}">${p}</button>`;
            if (p2 < totalPages) ph += `<span class="bt-pager-info">…</span><button data-pg="${totalPages}">${totalPages}</button>`;
            ph += `<button ${btState.page === totalPages ? 'disabled' : ''} data-pg="${btState.page + 1}">下一页 ›</button>`;
            pager.innerHTML = ph;
            pager.querySelectorAll('button[data-pg]').forEach(b => {
                b.addEventListener('click', () => { btState.page = parseInt(b.dataset.pg, 10); renderBankTable(); document.getElementById('banktable-tab').scrollIntoView({ behavior: 'instant' }); });
            });
        }

        function debounce(fn, ms) {
            let t = null;
            return function () {
                clearTimeout(t);
                t = setTimeout(fn, ms);
            };
        }

        /* ================================================================
         * 练习池管理（筛选 / 套题 / 错题重练）
         * ================================================================ */
        function resetToBankPool() {
            const filtered = getFilteredQuestions();
            pool = { list: filtered, mode: 'bank', label: '' };
            setSession = null;
            currentQuestionIndex = 0;
            updateResultCount(filtered.length);
            renderModeBanner();
            if (pool.list.length > 0) {
                showQuestionArea();
                displayQuestion();
            } else {
                showNoQuestion();
            }
        }

        function setPool(list, mode, label) {
            pool = { list: list, mode: mode, label: label };
            currentQuestionIndex = 0;
            renderModeBanner();
            if (pool.list.length > 0) {
                showQuestionArea();
                displayQuestion();
            } else {
                showNoQuestion();
            }
        }

        function exitSpecialMode() {
            if (setSession) saveSetInProgress();
            resetToBankPool();
            switchTab('practice');
        }

        function showQuestionArea() {
            uploadSection.style.display = 'none';
            questionContainer.style.display = 'block';
            emptyState.style.display = 'none';
            document.getElementById('filterToolbar').style.display = 'flex';
        }

        function showNoQuestion() {
            questionContainer.style.display = 'none';
            emptyState.style.display = 'block';
            uploadSection.style.display = 'none';
        }

        function renderModeBanner() {
            const banner = document.getElementById('modeBanner');
            const dots = document.getElementById('setDots');
            if (pool.mode === 'bank') {
                banner.classList.remove('show');
                dots.innerHTML = '';
                return;
            }
            banner.classList.add('show');
            document.getElementById('bannerText').textContent = pool.label;
            document.getElementById('bannerSub').textContent = `共 ${pool.list.length} 题`;
            if (setSession) {
                renderSetDots();
            } else {
                dots.innerHTML = '';
            }
        }

        function renderSetDots() {
            const dots = document.getElementById('setDots');
            if (!setSession) { dots.innerHTML = ''; return; }
            dots.innerHTML = setSession.list.map((q, i) => {
                const st = setSession.answers[q.id];
                let cls = 'sd-dot';
                if (st === true) cls += ' correct';
                else if (st === false) cls += ' wrong';
                if (i === currentQuestionIndex) cls += ' current';
                return `<span class="${cls}" onclick="jumpToPoolIndex(${i})">${i + 1}</span>`;
            }).join('');
        }

        window.jumpToPoolIndex = function (i) {
            if (i < 0 || i >= pool.list.length) return;
            currentQuestionIndex = i;
            displayQuestion();
        };

        /* ================================================================
         * 多维度筛选
         * ================================================================ */
        function buildFilterOptions() {
            const vals = (key) => [...new Set(questions.map(q => q[key]).filter(Boolean))].sort();
            fillSelect('fScenario', vals('scenario'), '全部场景');
            fillSelect('fPolarity', vals('polarity'), '肯定/否定');
            fillSelect('fFunction', vals('sentenceFunction'), '全部句子功能');
            fillSelect('fStructure', vals('sentenceStructure'), '全部句子结构');
            fillSelect('fPurpose', vals('purpose'), '全部交际目的');
        }

        function fillSelect(id, values, allLabel) {
            const sel = document.getElementById(id);
            const cur = sel.value;
            sel.innerHTML = `<option value="">${allLabel}</option>` +
                values.map(v => `<option value="${escapeHtml(v)}">${escapeHtml(v)}</option>`).join('');
            if (values.includes(cur)) sel.value = cur;
        }

        function getFilteredQuestions() {
            const sc = document.getElementById('fScenario').value;
            const po = document.getElementById('fPolarity').value;
            const fn = document.getElementById('fFunction').value;
            const st = document.getElementById('fStructure').value;
            const pu = document.getElementById('fPurpose').value;
            const kw = document.getElementById('fKeyword').value.trim().toLowerCase();
            return questions.filter(q => {
                if (sc && (q.scenario || '未分类') !== sc) return false;
                if (po && (q.polarity || '未分类') !== po) return false;
                if (fn && (q.sentenceFunction || '未分类') !== fn) return false;
                if (st && (q.sentenceStructure || '未分类') !== st) return false;
                if (pu && (q.purpose || '未分类') !== pu) return false;
                if (kw) {
                    const hay = (q.prompt + ' ' + q.questionText + ' ' + q.correctAnswer + ' ' + (q.source || '')).toLowerCase();
                    if (!hay.includes(kw)) return false;
                }
                return true;
            });
        }

        function applyFilters() {
            // 筛选仅在题库练习模式下生效；套题/错题重练中忽略
            if (pool.mode !== 'bank') {
                showToast('当前处于「' + pool.label + '」模式，退出后才能使用筛选');
                return;
            }
            resetToBankPool();
        }

        function resetAllFilters() {
            ['fScenario', 'fPolarity', 'fFunction', 'fStructure', 'fPurpose'].forEach(id => {
                document.getElementById(id).value = '';
            });
            document.getElementById('fKeyword').value = '';
            if (pool.mode !== 'bank') { exitSpecialMode(); return; }
            resetToBankPool();
        }

        function updateResultCount(n) {
            document.getElementById('resultCount').textContent = `共 ${n} 题`;
        }

        /* ================================================================
         * 套题模块
         * ================================================================ */
        function deriveSetId(source, id, date) {
            const s = (source || '').trim();
            if (s && s !== 'None') return s;
            const m = (id || '').match(/^(.*?)_W_/);
            if (m && m[1]) return m[1];
            if (date) return String(date).slice(0, 10);
            return '未分组';
        }

        function getSets() {
            const map = new Map();   // setId -> {id, date, questions[]}
            questions.forEach(q => {
                const sid = q.source || deriveSetId(q.source, q.id, q.date);
                if (!map.has(sid)) map.set(sid, { id: sid, date: q.date || '', questions: [] });
                const set = map.get(sid);
                set.questions.push(q);
                if (!set.date && q.date) set.date = q.date;
            });
            const list = Array.from(map.values());
            list.sort((a, b) => (a.date || '9999').localeCompare(b.date || '9999') || a.id.localeCompare(b.id));
            return list;
        }

        function loadSetStats() {
            try {
                setStats = JSON.parse(localStorage.getItem(SETS_KEY) || '{}') || {};
            } catch (e) { setStats = {}; }
        }

        function saveSetStats() {
            try {
                localStorage.setItem(SETS_KEY, JSON.stringify(setStats));
            } catch (e) { /* ignore */ }
        }

        function saveSetInProgress() {
            if (!setSession) return;
            const st = setStats[setSession.setId] || (setStats[setSession.setId] = { attempts: [], inProgress: {} });
            st.inProgress = Object.assign({}, setSession.answers);
            saveSetStats();
        }

        function renderSets() {
            const sets = getSets();
            const grid = document.getElementById('setsGrid');
            const empty = document.getElementById('setsEmpty');
            if (sets.length === 0) {
                grid.innerHTML = '';
                empty.style.display = 'block';
                return;
            }
            empty.style.display = 'none';

            let doneCount = 0;
            const cards = sets.map((set, idx) => {
                const n = idx + 1;
                const st = setStats[set.id];
                const total = set.questions.length;
                const answered = st && st.inProgress ? Object.keys(st.inProgress).filter(k => set.questions.some(q => q.id === k)).length : 0;
                const attempts = (st && st.attempts) ? st.attempts : [];
                const best = attempts.length ? Math.max(...attempts.map(a => a.correct)) : null;
                const last = attempts.length ? attempts[attempts.length - 1] : null;
                const completed = attempts.length > 0;
                if (completed) doneCount++;

                let statusPill, btnLabel;
                if (completed && answered === 0) { statusPill = '<span class="set-status-pill sp-done">✓ 已完成</span>'; btnLabel = '再练一遍'; }
                else if (answered > 0) { statusPill = `<span class="set-status-pill sp-doing">⏳ 进行中 ${answered}/${total}</span>`; btnLabel = '继续练习'; }
                else { statusPill = '<span class="set-status-pill sp-undo">未开始</span>'; btnLabel = completed ? '再练一遍' : '开始练习'; }

                const scoresHTML = attempts.length ? `
                    <div class="set-scores">
                        练习 ${attempts.length} 次 · 最近 ${last.correct}/${last.total}
                        ${best !== null ? ` · 最好 <strong style="color:var(--primary-color);">${best}/${total}</strong>` : ''}
                    </div>` : '';

                return `
                    <div class="set-card">
                        <div class="set-head">
                            <span class="set-no">第 ${n} 套</span>
                            <span class="set-date">${escapeHtml(set.date || '')}</span>
                        </div>
                        <div class="set-name">${escapeHtml(set.id)}</div>
                        <div class="set-meta">${total} 题 · ${statusPill}</div>
                        ${scoresHTML}
                        <button class="btn btn-primary" onclick="startSet('${escapeJs(set.id)}', ${completed && answered === 0})">${btnLabel}</button>
                    </div>`;
            }).join('');

            grid.innerHTML = cards;
            document.getElementById('setsSummaryLine').textContent =
                `共 ${sets.length} 套 · ${questions.length} 题 · 已完成 ${doneCount} 套`;
        }

        window.startSet = function (setId, restart) {
            const set = getSets().find(s => s.id === setId);
            if (!set || set.questions.length === 0) { showToast('未找到该套题'); return; }

            let answers = {};
            const st = setStats[setId];
            if (!restart && st && st.inProgress) {
                answers = Object.assign({}, st.inProgress);
                // 只保留当前仍存在的题目
                Object.keys(answers).forEach(qid => {
                    if (!set.questions.some(q => q.id === qid)) delete answers[qid];
                });
                if (Object.keys(answers).length === set.questions.length) {
                    // 旧的进行中记录其实已完成 → 当作重新开始
                    answers = {};
                }
            }

            setSession = { setId: setId, list: set.questions, answers: answers };
            setPool(set.questions, 'set', `📚 套题 · ${setId}`);
            switchTab('practice');
            showToast(`${setId}：共 ${set.questions.length} 题，加油！`);
        };

        function recordSetAnswer(q, isCorrect) {
            if (!setSession || setSession.list.indexOf(q) === -1) return;
            setSession.answers[q.id] = isCorrect;
            saveSetInProgress();
            renderSetDots();
            if (Object.keys(setSession.answers).length >= setSession.list.length) {
                setTimeout(completeSet, 600);
            }
        }

        function completeSet() {
            if (!setSession) return;
            const total = setSession.list.length;
            const correct = setSession.list.filter(q => setSession.answers[q.id] === true).length;
            const st = setStats[setSession.setId] || (setStats[setSession.setId] = { attempts: [], inProgress: {} });
            st.attempts.push({ at: new Date().toISOString(), correct: correct, total: total });
            st.inProgress = {};
            saveSetStats();
            showSetSummary(setSession.setId, correct, total);
        }

        function showSetSummary(setId, correct, total) {
            const pct = total > 0 ? Math.round(correct / total * 100) : 0;
            document.getElementById('summaryTitle').textContent = '🎉 套题完成 · ' + setId;
            document.getElementById('summaryScore').textContent = `${correct} / ${total}`;
            document.getElementById('summaryPct').textContent = `${pct}% 正确率`;

            const session = setSession;
            const wrongs = session ? session.list.filter(q => session.answers[q.id] === false) : [];
            document.getElementById('summaryDetail').innerHTML = wrongs.length === 0
                ? '<p style="text-align:center; color:#2e7d32; font-weight:600;">满分通过，太强了！💪</p>'
                : wrongs.map(q => `
                    <div class="summary-wrong-item">
                        <div>❓ ${escapeHtml(q.prompt)}</div>
                        <div>✅ <strong>${escapeHtml(q.correctAnswer)}</strong></div>
                    </div>`).join('');

            document.getElementById('summaryRestart').dataset.setId = setId;
            document.getElementById('summaryWrong').dataset.wrong = JSON.stringify(wrongs.map(q => q.id));
            document.getElementById('summaryWrong').style.display = wrongs.length ? '' : 'none';
            document.getElementById('setSummaryModal').classList.add('show');
        }

        function closeSetSummary() {
            document.getElementById('setSummaryModal').classList.remove('show');
        }

        function startReviewPractice(questionIds, label) {
            const list = questionIds
                .map(id => questions.find(q => q.id === id))
                .filter(Boolean);
            if (list.length === 0) { showToast('没有可重练的题目'); return; }
            setSession = null;
            setPool(list, 'review', '📌 ' + label);
            switchTab('practice');
        }

        /* ================================================================
         * 文件上传（兼容新版 625 表格与旧版表格两种列结构）
         * ================================================================ */
        function reuploadFile() {
            if (confirm('导入自定义题库将替换当前所有题目（练习记录与错题本保留），确定继续吗？')) {
                fileInput.click();
            }
        }

        function handleDragOver(e) {
            e.preventDefault();
            uploadSection.classList.add('dragover');
        }

        function handleDragLeave(e) {
            e.preventDefault();
            uploadSection.classList.remove('dragover');
        }

        function handleDrop(e) {
            e.preventDefault();
            uploadSection.classList.remove('dragover');
            const file = e.dataTransfer.files[0];
            if (file && (file.name.endsWith('.xlsx') || file.name.endsWith('.xls'))) {
                processFile(file);
            }
        }

        function handleFileSelect(e) {
            const file = e.target.files[0];
            if (file) processFile(file);
            e.target.value = '';
        }

        function processFile(file) {
            const reader = new FileReader();
            reader.onload = function (e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                const count = parseExcelData(jsonData);
                if (count > 0) {
                    persistBank();
                    buildFilterOptions();
                    applyFiltersResetUI();
                    showToast(`成功加载 ${count} 道题目！`);
                } else {
                    showToast('未解析到有效题目，请检查文件格式');
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function applyFiltersResetUI() {
            ['fScenario', 'fPolarity', 'fFunction', 'fStructure', 'fPurpose'].forEach(id => {
                document.getElementById(id).value = '';
            });
            document.getElementById('fKeyword').value = '';
            setSession = null;
            pool = { list: questions.slice(), mode: 'bank', label: '' };
            currentQuestionIndex = 0;
            updateResultCount(questions.length);
            renderModeBanner();
            if (pool.list.length > 0) displayQuestion();
        }

        function parseExcelData(data) {
            if (!data || data.length < 2) return 0;
            const headers = (data[0] || []).map(h => String(h == null ? '' : h).trim());
            let fmt = 'legacy';
            const idIdx = headers.indexOf('编号');
            if (idIdx === 4 || headers.includes('场景') || headers.includes('交际目的')) fmt = 'new';
            else if (idIdx === 7) fmt = 'legacy';

            const S = v => (v == null ? '' : String(v).trim());
            const parsed = [];
            const seenIds = {};

            for (let i = 1; i < data.length; i++) {
                const row = data[i];
                if (!row) continue;

                let q;
                if (fmt === 'new') {
                    q = {
                        id: S(row[4]),
                        source: S(row[2]),
                        date: S(row[1]).slice(0, 10),
                        instruction: S(row[5]),
                        prompt: S(row[6]),
                        questionText: S(row[7]),
                        options: [],
                        answerOrder: [],
                        correctAnswer: S(row[23]),
                        explanation: S(row[24]),
                        scenario: S(row[25]),
                        polarity: S(row[27]),
                        sentenceFunction: S(row[28]),
                        sentenceStructure: S(row[29]),
                        grammar: S(row[30]),
                        purpose: S(row[31])
                    };
                    for (let j = 8; j <= 15; j++) if (S(row[j])) q.options.push(S(row[j]));
                    for (let j = 16; j <= 22; j++) {
                        const L = S(row[j]);
                        if (L && L[0] >= 'A' && L[0] <= 'H') {
                            const oi = L.charCodeAt(0) - 65;
                            if (q.options[oi]) q.answerOrder.push(q.options[oi]);
                        }
                    }
                } else {
                    q = {
                        id: S(row[7]),
                        source: S(row[3]),
                        date: S(row[4]).slice(0, 10),
                        instruction: S(row[9]),
                        prompt: S(row[10]),
                        questionText: S(row[14]),
                        options: [],
                        answerOrder: [],
                        correctAnswer: S(row[33]),
                        explanation: S(row[34]),
                        scenario: '',
                        polarity: '',
                        sentenceFunction: '',
                        sentenceStructure: '',
                        grammar: '',
                        purpose: ''
                    };
                    for (let j = 15; j <= 22; j++) if (S(row[j])) q.options.push(S(row[j]));
                    for (let j = 23; j <= 32; j++) {
                        const L = S(row[j]);
                        if (L && L[0] >= 'A' && L[0] <= 'H') {
                            const oi = L.charCodeAt(0) - 65;
                            if (q.options[oi]) q.answerOrder.push(q.options[oi]);
                        }
                    }
                }

                if (!q.prompt && !q.questionText) continue;
                if (q.options.length === 0) continue;
                if (q.options.every(o => o === '选项缺失')) continue;

                if (!q.id) q.id = 'q' + i;
                if (seenIds[q.id]) {
                    seenIds[q.id]++;
                    q.id = q.id + '#' + seenIds[q.id];
                } else {
                    seenIds[q.id] = 1;
                }

                q.source = q.source || deriveSetId('', q.id, q.date);
                parsed.push(q);
            }

            questions = parsed;
            return parsed.length;
        }

        /* ================================================================
         * 题目渲染（槽位解析兼容 1~13 个下划线的任意占位长度）
         * ================================================================ */
        function parseQuestionTemplate(template) {
            const parts = [];
            let slotIndex = 0;
            const segments = template.split(/_+/);
            segments.forEach((segment, index) => {
                if (segment) {
                    const words = segment.trim().split(/(\s+|[？?!.,;:])/).filter(s => s.trim());
                    words.forEach(word => {
                        if (word.trim()) {
                            parts.push({ type: 'word', content: word });
                        }
                    });
                }
                if (index < segments.length - 1 || template.endsWith('_')) {
                    parts.push({ type: 'slot', index: slotIndex++ });
                }
            });
            return parts;
        }

        function countTemplateSlots(template) {
            const m = template.match(/_+/g);
            return m ? m.length : 0;
        }

        function displayQuestion() {
            if (pool.list.length === 0) return;
            if (currentQuestionIndex >= pool.list.length) currentQuestionIndex = 0;

            currentQuestion = pool.list[currentQuestionIndex];
            slots = new Array(countTemplateSlots(currentQuestion.questionText)).fill(null);
            usedWords.clear();

            const parts = parseQuestionTemplate(currentQuestion.questionText);
            currentQuestion.firstSlotIsSentenceStart = parts.length > 0 && parts[0].type === 'slot';

            promptText.textContent = currentQuestion.prompt || '（无背景材料）';
            currentNum.textContent = currentQuestionIndex + 1;
            totalNum.textContent = pool.list.length;
            progressFill.style.width = `${((currentQuestionIndex + 1) / pool.list.length) * 100}%`;

            renderSentence();
            renderOptions();
            renderSetDots();

            feedback.classList.remove('show');
            document.getElementById('checkBtn').disabled = false;
        }

        function renderSentence() {
            const parts = parseQuestionTemplate(currentQuestion.questionText);
            sentenceContainer.innerHTML = '';

            parts.forEach(part => {
                if (part.type === 'word') {
                    const span = document.createElement('span');
                    span.className = 'sentence-word';
                    span.textContent = part.content + ' ';
                    sentenceContainer.appendChild(span);
                } else if (part.type === 'slot') {
                    const slot = document.createElement('div');
                    slot.className = 'slot';
                    slot.dataset.slotIndex = part.index;
                    slot.innerHTML = `<span class="slot-slot-number">${part.index + 1}</span>_____`;

                    slot.addEventListener('click', () => handleSlotClick(part.index));

                    slot.addEventListener('dragover', (e) => {
                        e.preventDefault();
                        e.dataTransfer.dropEffect = 'move';
                        slot.classList.add('drag-over');
                    });

                    slot.addEventListener('dragleave', () => {
                        slot.classList.remove('drag-over');
                    });

                    slot.addEventListener('drop', (e) => {
                        e.preventDefault();
                        slot.classList.remove('drag-over');
                        const word = e.dataTransfer.getData('text/plain');
                        if (word && !usedWords.has(word)) {
                            fillSlot(part.index, word);
                        }
                    });

                    sentenceContainer.appendChild(slot);
                }
            });
        }

        function renderOptions() {
            optionsArea.innerHTML = '';
            const options = currentQuestion.options;

            options.forEach((option, index) => {
                const block = document.createElement('div');
                block.className = 'word-block';
                const letter = String.fromCharCode(65 + index);
                block.innerHTML = `<span class="opt-letter">${letter}</span><span>${escapeHtml(option)}</span>`;
                block.dataset.word = option;
                block.draggable = true;

                if (usedWords.has(option)) {
                    block.classList.add('used');
                }

                block.addEventListener('click', () => handleWordClick(option));

                block.addEventListener('dragstart', (e) => {
                    if (usedWords.has(option)) {
                        e.preventDefault();
                        return;
                    }
                    e.dataTransfer.setData('text/plain', option);
                    e.dataTransfer.effectAllowed = 'move';
                    block.classList.add('dragging');
                });

                block.addEventListener('dragend', () => {
                    block.classList.remove('dragging');
                });

                optionsArea.appendChild(block);
            });
        }

        function handleWordClick(word) {
            if (usedWords.has(word)) return;
            const emptySlotIndex = slots.findIndex(s => s === null);
            if (emptySlotIndex === -1) {
                showToast('所有空位已填满，请点击空位修改');
                return;
            }
            fillSlot(emptySlotIndex, word);
        }

        function handleSlotClick(slotIndex) {
            const currentWord = slots[slotIndex];
            if (currentWord) {
                slots[slotIndex] = null;
                usedWords.delete(currentWord);
                updateSlotDisplay(slotIndex, null);
                updateWordDisplay(currentWord, false);
            }
        }

        function fillSlot(slotIndex, word) {
            const previousSlotIndex = slots.indexOf(word);
            if (previousSlotIndex !== -1) {
                slots[previousSlotIndex] = null;
                updateSlotDisplay(previousSlotIndex, null);
            }
            slots[slotIndex] = word;
            usedWords.add(word);
            updateSlotDisplay(slotIndex, word);
            updateWordDisplay(word, true);
        }

        function updateSlotDisplay(slotIndex, word) {
            const slotElement = sentenceContainer.querySelector(`[data-slot-index="${slotIndex}"]`);
            if (!slotElement) return;

            if (word) {
                slotElement.classList.add('filled');
                let displayWord = word;
                if (slotIndex === 0 && currentQuestion.firstSlotIsSentenceStart) {
                    displayWord = word.charAt(0).toUpperCase() + word.slice(1);
                }
                slotElement.innerHTML = `<span class="slot-slot-number">${slotIndex + 1}</span>${escapeHtml(displayWord)}`;
            } else {
                slotElement.classList.remove('filled', 'correct', 'incorrect');
                slotElement.innerHTML = `<span class="slot-slot-number">${slotIndex + 1}</span>_____`;
            }
        }

        function updateWordDisplay(word, isUsed) {
            const wordBlock = Array.from(optionsArea.querySelectorAll('.word-block'))
                .find(block => block.dataset.word === word);
            if (wordBlock) {
                if (isUsed) wordBlock.classList.add('used');
                else wordBlock.classList.remove('used');
            }
        }

        function resetQuestion() {
            slots.fill(null);
            usedWords.clear();
            displayQuestion();
        }

        function prevQuestion() {
            if (currentQuestionIndex > 0) {
                currentQuestionIndex--;
                displayQuestion();
            }
        }

        function nextQuestion() {
            if (currentQuestionIndex < pool.list.length - 1) {
                currentQuestionIndex++;
                displayQuestion();
            } else if (setSession && Object.keys(setSession.answers).length >= setSession.list.length) {
                completeSet();
            }
        }

        /* ================================================================
         * 检查答案 + 错题自动收录 + 套题进度记录
         * ================================================================ */
        function checkAnswer() {
            const parts = parseQuestionTemplate(currentQuestion.questionText);
            let sentenceParts = [];
            let slotPointer = 0;
            let shouldCapitalize = currentQuestion.firstSlotIsSentenceStart;

            parts.forEach((part) => {
                if (part.type === 'word') {
                    sentenceParts.push(part.content);
                    shouldCapitalize = false;
                } else if (part.type === 'slot') {
                    if (slots[slotPointer]) {
                        let word = slots[slotPointer];
                        if (shouldCapitalize && slotPointer === 0) {
                            word = word.charAt(0).toUpperCase() + word.slice(1);
                            shouldCapitalize = false;
                        }
                        sentenceParts.push(word);
                    }
                    slotPointer++;
                }
            });

            const userSentence = sentenceParts.join(' ');

            const getWords = (str) => {
                return str
                    .replace(/[’‘]/g, "'")
                    .replace(/[？?!,.;:]/g, ' ')
                    .split(/\s+/)
                    .filter(w => w.length > 0)
                    .map(w => w.toLowerCase())
                    // 单独的小写 l 几乎必是 I 的录入错误，统一按 i 处理
                    .map(w => w === 'l' ? 'i' : w);
            };

            const userWordsLower = getWords(userSentence);
            const correctWords = getWords(currentQuestion.correctAnswer || '');

            const isCorrect = userWordsLower.length === correctWords.length &&
                userWordsLower.every((word, index) => word === correctWords[index]);

            feedback.classList.add('show');
            feedback.classList.remove('correct', 'incorrect');

            const dimTagsHTML = buildDimTagsHTML(currentQuestion);

            if (isCorrect) {
                feedback.classList.add('correct');
                feedback.innerHTML = `
                    <strong>✓ 正确！</strong>
                    <p>答案：${escapeHtml(currentQuestion.correctAnswer)}</p>
                    ${currentQuestion.explanation ? `<p style="font-size:0.9em; opacity:0.85;">${escapeHtml(currentQuestion.explanation)}</p>` : ''}
                    ${dimTagsHTML}
                `;
                stats.correct++;
                stats.streak++;
                markSlotsAsCorrect();
            } else {
                feedback.classList.add('incorrect');
                feedback.innerHTML = `
                    <strong>✗ 不正确</strong>
                    <p>你的答案：${escapeHtml(userSentence)}</p>
                    <p>正确答案：<strong>${escapeHtml(currentQuestion.correctAnswer)}</strong></p>
                    ${currentQuestion.explanation ? `<p style="font-size:0.9em; opacity:0.85;">${escapeHtml(currentQuestion.explanation)}</p>` : ''}
                    ${dimTagsHTML}
                `;
                stats.streak = 0;
                markSlotsAsIncorrect();
                addToWrongBook(currentQuestion, userSentence);
            }

            stats.total++;
            stats.history.unshift({
                questionId: currentQuestion.id,
                prompt: (currentQuestion.prompt || '').substring(0, 50) + '...',
                userAnswer: userSentence,
                correctAnswer: currentQuestion.correctAnswer,
                isCorrect: isCorrect,
                timestamp: new Date().toISOString()
            });

            if (stats.history.length > 100) {
                stats.history = stats.history.slice(0, 100);
            }

            saveStats();
            recordSetAnswer(currentQuestion, isCorrect);
            document.getElementById('checkBtn').disabled = true;
        }

        function buildDimTagsHTML(q) {
            const dims = [
                ['场景', q.scenario], ['肯/否', q.polarity], ['功能', q.sentenceFunction],
                ['结构', q.sentenceStructure], ['目的', q.purpose]
            ].filter(d => d[1]);
            if (dims.length === 0) return '';
            return `<div class="dim-tags">${dims.map(d =>
                `<span class="dim-tag">${d[0]}：${escapeHtml(d[1])}</span>`).join('')}</div>`;
        }

        function markSlotsAsCorrect() {
            sentenceContainer.querySelectorAll('.slot.filled').forEach(slot => {
                slot.classList.add('correct');
            });
        }

        function markSlotsAsIncorrect() {
            sentenceContainer.querySelectorAll('.slot.filled').forEach(slot => {
                slot.classList.add('incorrect');
            });
        }

        /* ================================================================
         * 错题本模块（参考 complete the words 设计）
         * ================================================================ */
        function wbId() {
            return 'wb-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 8);
        }

        function wbToday() {
            const d = new Date();
            return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
        }

        function loadWrongBook() {
            try {
                const raw = JSON.parse(localStorage.getItem(WB_KEY) || 'null');
                if (raw && Array.isArray(raw.items)) {
                    wrongBookItems = raw.items.filter(it => it && it.id && it.questionId);
                }
            } catch (e) { wrongBookItems = []; }
        }

        function saveWrongBook() {
            try {
                localStorage.setItem(WB_KEY, JSON.stringify({
                    schemaVersion: '1.0',
                    savedAt: new Date().toISOString(),
                    items: wrongBookItems
                }));
                return true;
            } catch (e) { return false; }
        }

        function addToWrongBook(q, userSentence) {
            const existing = wrongBookItems.find(it => it.questionId === q.id);
            if (existing) {
                existing.wrongCount = (existing.wrongCount || 1) + 1;
                existing.myAnswer = userSentence;
                existing.updatedAt = new Date().toISOString();
                if (existing.status === '已掌握') existing.status = '复习中';
            } else {
                wrongBookItems.push({
                    id: wbId(),
                    questionId: q.id,
                    setId: q.source || '',
                    prompt: q.prompt || '',
                    questionText: q.questionText || '',
                    myAnswer: userSentence,
                    correctAnswer: q.correctAnswer || '',
                    explanation: q.explanation || '',
                    scenario: q.scenario || '未分类',
                    polarity: q.polarity || '',
                    sentenceFunction: q.sentenceFunction || '未分类',
                    sentenceStructure: q.sentenceStructure || '未分类',
                    purpose: q.purpose || '未分类',
                    reasonCategory: '',
                    reasonNote: '',
                    date: wbToday(),
                    status: '未掌握',
                    wrongCount: 1,
                    createdAt: new Date().toISOString()
                });
            }
            saveWrongBook();
            const fb = document.getElementById('feedback');
            if (fb) {
                fb.innerHTML += `<p style="margin-top:8px; font-size:0.85em;">📥 已自动收录到错题本（可在「错题本」页管理）</p>`;
            }
        }

        function renderWrongBook() {
            renderWbDashboard();
            renderWbFilterOptions();
            renderWbList();
        }

        function renderWbDashboard() {
            const items = wrongBookItems;
            const byStatus = s => items.filter(it => it.status === s).length;
            document.getElementById('wbDashboard').innerHTML = `
                <div class="wb-stat-card"><div class="num">${items.length}</div><div class="lbl">错题总数</div></div>
                <div class="wb-stat-card st-undo"><div class="num">${byStatus('未掌握')}</div><div class="lbl">🔴 未掌握</div></div>
                <div class="wb-stat-card st-doing"><div class="num">${byStatus('复习中')}</div><div class="lbl">🟡 复习中</div></div>
                <div class="wb-stat-card st-done"><div class="num">${byStatus('已掌握')}</div><div class="lbl">🟢 已掌握</div></div>`;

            // 场景占比
            const scCount = {};
            items.forEach(it => { const d = it.scenario || '未分类'; scCount[d] = (scCount[d] || 0) + 1; });
            const scMax = Math.max(1, ...Object.values(scCount));
            const scHTML = Object.entries(scCount).sort((a, b) => b[1] - a[1]).slice(0, 8).map(([d, n]) => `
                <div class="wb-bar-row">
                    <span class="bar-label">${escapeHtml(d)}</span>
                    <div class="bar-track"><div class="bar-fill" style="width:${Math.round(n / scMax * 100)}%"></div></div>
                    <span class="bar-val">${n}</span>
                </div>`).join('');
            document.getElementById('wbScenarioChart').innerHTML = scHTML || '<div style="color:#aaa;font-size:0.85em;">暂无数据</div>';

            // 错误原因 Top5
            const rCount = {};
            items.forEach(it => { if (it.reasonCategory) rCount[it.reasonCategory] = (rCount[it.reasonCategory] || 0) + 1; });
            const rMax = Math.max(1, ...Object.values(rCount));
            const rHTML = Object.entries(rCount).sort((a, b) => b[1] - a[1]).slice(0, 5).map(([r, n]) => `
                <div class="wb-bar-row">
                    <span class="bar-label">${escapeHtml(r)}</span>
                    <div class="bar-track"><div class="bar-fill" style="width:${Math.round(n / rMax * 100)}%; background:#E8A317;"></div></div>
                    <span class="bar-val">${n}</span>
                </div>`).join('');
            document.getElementById('wbReasonChart').innerHTML = rHTML || '<div style="color:#aaa;font-size:0.85em;">在错题上标注错误原因后此处汇总</div>';

            // 句子结构词云
            const stCount = {};
            items.forEach(it => { if (it.sentenceStructure) stCount[it.sentenceStructure] = (stCount[it.sentenceStructure] || 0) + 1; });
            const stMax = Math.max(1, ...Object.values(stCount));
            const stHTML = Object.entries(stCount).sort((a, b) => b[1] - a[1]).map(([k, n]) => {
                const f = n / stMax > 0.6 ? 'f3' : (n / stMax > 0.3 ? 'f2' : '');
                return `<span class="wb-kp-tag ${f}">${escapeHtml(k)} ×${n}</span>`;
            }).join('');
            document.getElementById('wbStructureCloud').innerHTML = stHTML || '<span style="color:#aaa;font-size:0.85em;">暂无数据</span>';
        }

        function renderWbFilterOptions() {
            const fill = (id, key, label) => {
                const sel = document.getElementById(id);
                const cur = sel.value;
                const vals = [...new Set(wrongBookItems.map(it => it[key]).filter(Boolean))].sort();
                sel.innerHTML = `<option value="">${label}</option>` +
                    vals.map(v => `<option value="${escapeHtml(v)}"${v === cur ? ' selected' : ''}>${escapeHtml(v)}</option>`).join('');
            };
            fill('wbScenarioFilter', 'scenario', '全部场景');
            fill('wbFunctionFilter', 'sentenceFunction', '全部句子功能');
            fill('wbPurposeFilter', 'purpose', '全部交际目的');
        }

        function wbMatches(it) {
            const sc = document.getElementById('wbScenarioFilter').value;
            const fn = document.getElementById('wbFunctionFilter').value;
            const pu = document.getElementById('wbPurposeFilter').value;
            const range = document.getElementById('wbDateFilter').value;
            const s = document.getElementById('wbStatusFilter').value;
            const kw = document.getElementById('wbSearch').value.trim().toLowerCase();

            if (sc && it.scenario !== sc) return false;
            if (fn && it.sentenceFunction !== fn) return false;
            if (pu && it.purpose !== pu) return false;
            if (s && it.status !== s) return false;
            if (!s && it.status === '已掌握') return false;
            if (range) {
                const limit = new Date(Date.now() - parseInt(range) * 86400000);
                if (new Date(it.date || '1970-01-01') < limit) return false;
            }
            if (kw) {
                const hay = (it.prompt + ' ' + it.questionText + ' ' + it.myAnswer + ' ' + it.correctAnswer + ' ' + (it.reasonNote || '')).toLowerCase();
                if (!hay.includes(kw)) return false;
            }
            return true;
        }

        function wbCardHTML(it) {
            const stIdx = WB_STATUSES.indexOf(it.status);
            const dims = [
                it.scenario, it.sentenceFunction, it.sentenceStructure, it.purpose
            ].filter(Boolean);
            const inBank = questions.some(q => q.id === it.questionId);
            return `
                <div class="wb-card st-${stIdx === 2 ? 'done' : (stIdx === 1 ? 'doing' : 'undo')}">
                    <div class="wb-card-head">
                        <span class="wb-qid">${escapeHtml(it.questionId)}${it.setId ? ' · ' + escapeHtml(it.setId) : ''}</span>
                        ${dims.map(d => `<span class="wb-badge">${escapeHtml(d)}</span>`).join('')}
                        <span class="wb-badge date">📅 ${escapeHtml(it.date || '')}</span>
                    </div>
                    <div class="wb-compare">
                        ❓ ${escapeHtml(it.prompt)}<br>
                        我造：${it.myAnswer ? '<span class="mine">' + escapeHtml(it.myAnswer) + '</span>' : '<span style="color:#999;">（未记录）</span>'}
                        　→　正确：<span class="right">${escapeHtml(it.correctAnswer)}</span>
                        ${it.wrongCount > 1 ? `<span class="wb-badge" style="margin-left:6px;">累计答错 ${it.wrongCount} 次</span>` : ''}
                    </div>
                    <div class="wb-reason">错误原因：<strong>${escapeHtml(it.reasonCategory || '未分类')}</strong>${it.reasonNote ? ' — ' + escapeHtml(it.reasonNote) : ''}</div>
                    <div class="wb-card-foot">
                        <div class="wb-status-group">
                            ${WB_STATUSES.map((st, i) => `<button class="${i === stIdx ? (i === 0 ? 'on-undo' : (i === 1 ? 'on-doing' : 'on-done')) : ''}" onclick="wbSetStatus('${it.id}', '${st}')">${['🔴 未掌握', '🟡 复习中', '🟢 已掌握'][i]}</button>`).join('')}
                        </div>
                        <span style="flex:1;"></span>
                        ${inBank ? `<button class="wb-mini-btn wb-jump-btn" onclick="wbJumpToQuestion('${it.id}')" title="跳转到原题">↩ 返回题目</button>` : ''}
                        <button class="wb-mini-btn" onclick="openWbEdit('${it.id}')">✏️ 标注错因</button>
                        <button class="wb-mini-btn" id="del-${it.id}" onclick="wbDeleteClick('${it.id}')">🗑️ 删除</button>
                    </div>
                </div>`;
        }

        function renderWbList() {
            const listEl = document.getElementById('wbList');
            const archiveEl = document.getElementById('wbArchive');
            const visible = wrongBookItems.filter(wbMatches);

            if (visible.length === 0) {
                listEl.innerHTML = `
                    <div class="wb-empty">
                        <div class="icon">${wrongBookItems.length === 0 ? '📕' : '🔍'}</div>
                        <h3>${wrongBookItems.length === 0 ? '错题本还是空的' : '没有符合条件的错题'}</h3>
                        <p>${wrongBookItems.length === 0 ? '去练习模式检查答案，答错的句子会自动收录进来' : '试试放宽筛选条件'}</p>
                    </div>`;
            } else {
                listEl.innerHTML = visible.map(wbCardHTML).join('');
            }

            const done = wrongBookItems.filter(it => it.status === '已掌握');
            if (done.length === 0) {
                archiveEl.innerHTML = '';
                return;
            }
            const open = archiveEl.dataset.open === '1';
            archiveEl.innerHTML = `
                <div class="wb-archive-head" onclick="toggleWbArchive()">
                    <span>${open ? '▾' : '▸'} 📦 已掌握归档（${done.length} 题）</span>
                    <span style="font-size:0.8em; font-weight:400; color:#8aa694;">点击${open ? '收起' : '展开'} · 可回退状态</span>
                </div>
                ${open ? done.map(wbCardHTML).join('') : ''}`;
        }

        window.toggleWbArchive = function () {
            const el = document.getElementById('wbArchive');
            el.dataset.open = el.dataset.open === '1' ? '0' : '1';
            renderWbList();
        };

        window.wbSetStatus = function (id, status) {
            const it = wrongBookItems.find(x => x.id === id);
            if (!it) return;
            it.status = status;
            it.updatedAt = new Date().toISOString();
            saveWrongBook();
            renderWrongBook();
        };

        window.wbDeleteClick = function (id) {
            const btn = document.getElementById('del-' + id);
            if (wbDeleteArmed === id) {
                wrongBookItems = wrongBookItems.filter(x => x.id !== id);
                clearTimeout(wbDeleteTimer);
                wbDeleteArmed = null;
                saveWrongBook();
                renderWrongBook();
                return;
            }
            if (wbDeleteArmed) {
                const oldBtn = document.getElementById('del-' + wbDeleteArmed);
                if (oldBtn) { oldBtn.classList.remove('danger-armed'); oldBtn.textContent = '🗑️ 删除'; }
                clearTimeout(wbDeleteTimer);
            }
            wbDeleteArmed = id;
            if (btn) { btn.classList.add('danger-armed'); btn.textContent = '⚠️ 再点一次确认'; }
            wbDeleteTimer = setTimeout(() => {
                if (wbDeleteArmed === id) {
                    wbDeleteArmed = null;
                    if (btn) { btn.classList.remove('danger-armed'); btn.textContent = '🗑️ 删除'; }
                }
            }, 3000);
        };

        function openWbEdit(id) {
            wbEditingId = id;
            const it = wrongBookItems.find(x => x.id === id);
            if (!it) return;
            document.getElementById('wbEditReason').innerHTML =
                '<option value="">（未分类）</option>' + WB_REASONS.map(r =>
                    `<option${it.reasonCategory === r ? ' selected' : ''}>${r}</option>`).join('');
            document.getElementById('wbEditStatus').value = it.status;
            document.getElementById('wbEditNote').value = it.reasonNote || '';
            document.getElementById('wbEditModal').classList.add('show');
        }

        function closeWbEdit() {
            document.getElementById('wbEditModal').classList.remove('show');
            wbEditingId = null;
        }

        function saveWbEdit() {
            const it = wrongBookItems.find(x => x.id === wbEditingId);
            if (!it) { closeWbEdit(); return; }
            it.reasonCategory = document.getElementById('wbEditReason').value;
            it.status = document.getElementById('wbEditStatus').value;
            it.reasonNote = document.getElementById('wbEditNote').value.trim();
            it.updatedAt = new Date().toISOString();
            saveWrongBook();
            closeWbEdit();
            renderWrongBook();
            showToast('错题信息已更新');
        }

        window.openWbEdit = openWbEdit;

        // 一键返回原题
        window.wbJumpToQuestion = function (id) {
            const it = wrongBookItems.find(x => x.id === id);
            if (!it || !it.questionId) { alert('该错题未关联题目，无法跳转。'); return; }
            const q = questions.find(x => x.id === it.questionId);
            if (!q) { alert('未在当前题库中找到对应题目（可能来自自定义导入），无法跳转。'); return; }

            // 切回题库练习模式并清除筛选，确保目标题在池中
            setSession = null;
            ['fScenario', 'fPolarity', 'fFunction', 'fStructure', 'fPurpose'].forEach(i => document.getElementById(i).value = '');
            document.getElementById('fKeyword').value = '';
            pool = { list: questions.slice(), mode: 'bank', label: '' };
            renderModeBanner();
            updateResultCount(questions.length);
            switchTab('practice');

            const idx = questions.findIndex(x => x.id === q.id);
            if (idx < 0) { alert('题目定位失败'); return; }
            currentQuestionIndex = idx;
            displayQuestion();

            // 滚动定位 + 高亮闪烁
            setTimeout(() => {
                const card = document.querySelector('.question-box');
                sentenceContainer.classList.remove('jump-flash');
                void sentenceContainer.offsetWidth;
                sentenceContainer.classList.add('jump-flash');
                if (card && card.scrollIntoView) {
                    try { card.scrollIntoView({ behavior: 'smooth', block: 'start' }); } catch (e) { /* ignore */ }
                }
            }, 100);
        };

        function startWrongbookPractice() {
            const targets = wrongBookItems
                .filter(it => it.status !== '已掌握')
                .map(it => questions.find(q => q.id === it.questionId))
                .filter(Boolean);
            if (targets.length === 0) {
                showToast(wrongBookItems.length === 0 ? '错题本是空的，先去练习吧' : '没有未掌握的错题，太棒了！');
                return;
            }
            setSession = null;
            setPool(targets, 'review', '📌 错题重练（未掌握）');
            switchTab('practice');
            showToast(`错题重练：共 ${targets.length} 题`);
        }

        function exportWrongBook() {
            if (wrongBookItems.length === 0) { alert('错题本为空，无需备份'); return; }
            const payload = {
                app: 'toefl-build-a-sentence',
                type: 'wrongbook-backup',
                schemaVersion: '1.0',
                exportedAt: new Date().toISOString(),
                items: wrongBookItems
            };
            const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = '句子构建错题本备份_' + wbToday() + '.json';
            a.click();
            URL.revokeObjectURL(a.href);
        }

        function importWrongBook(e) {
            const input = e.target;
            const file = input.files[0];
            input.value = '';
            if (!file) return;
            const reader = new FileReader();
            reader.onload = () => {
                try {
                    const data = JSON.parse(reader.result);
                    const items = Array.isArray(data) ? data : (Array.isArray(data.items) ? data.items : null);
                    if (!items) { alert('文件格式不正确：未找到错题数据'); return; }
                    let added = 0, skipped = 0;
                    items.forEach(raw => {
                        if (!raw || !raw.id || !raw.questionId) { skipped++; return; }
                        if (wrongBookItems.some(it => it.id === raw.id)) { skipped++; return; }
                        wrongBookItems.push({
                            id: raw.id,
                            questionId: raw.questionId,
                            setId: raw.setId || '',
                            prompt: raw.prompt || '',
                            questionText: raw.questionText || '',
                            myAnswer: raw.myAnswer || '',
                            correctAnswer: raw.correctAnswer || '',
                            explanation: raw.explanation || '',
                            scenario: raw.scenario || '未分类',
                            polarity: raw.polarity || '',
                            sentenceFunction: raw.sentenceFunction || '未分类',
                            sentenceStructure: raw.sentenceStructure || '未分类',
                            purpose: raw.purpose || '未分类',
                            reasonCategory: raw.reasonCategory || '',
                            reasonNote: raw.reasonNote || '',
                            date: raw.date || wbToday(),
                            status: WB_STATUSES.includes(raw.status) ? raw.status : '未掌握',
                            wrongCount: raw.wrongCount || 1,
                            createdAt: raw.createdAt || new Date().toISOString()
                        });
                        added++;
                    });
                    saveWrongBook();
                    renderWrongBook();
                    alert(`导入完成：新增 ${added} 条${skipped ? '，跳过 ' + skipped + ' 条（重复或无效）' : ''}`);
                } catch (err) {
                    alert('导入失败：文件不是有效的 JSON');
                }
            };
            reader.readAsText(file);
        }

        /* ================================================================
         * 统计
         * ================================================================ */
        function loadStats() {
            const saved = localStorage.getItem('toeflBuildStats');
            if (saved) {
                try {
                    stats = JSON.parse(saved);
                    if (!Array.isArray(stats.history)) stats.history = [];
                } catch (e) { /* ignore */ }
            }
        }

        function saveStats() {
            localStorage.setItem('toeflBuildStats', JSON.stringify(stats));
        }

        function clearStats() {
            if (confirm('确定要清除所有统计数据吗？（错题本不受影响）')) {
                stats = { total: 0, correct: 0, streak: 0, history: [] };
                saveStats();
                updateStatsDisplay();
                showToast('统计数据已清除');
            }
        }

        function exportStats() {
            const dataStr = JSON.stringify(stats, null, 2);
            const blob = new Blob([dataStr], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `toefl-build-sentence-stats-${new Date().toISOString().split('T')[0]}.json`;
            a.click();
            showToast('统计数据已导出');
        }

        function updateStatsDisplay() {
            document.getElementById('statTotal').textContent = stats.total;
            document.getElementById('statCorrect').textContent = stats.correct;
            document.getElementById('statRate').textContent = stats.total > 0
                ? Math.round((stats.correct / stats.total) * 100) + '%'
                : '0%';
            document.getElementById('statStreak').textContent = stats.streak;
            const completedSets = Object.values(setStats).filter(st => st.attempts && st.attempts.length > 0).length;
            document.getElementById('statSets').textContent = completedSets;
            document.getElementById('statWrong').textContent = wrongBookItems.filter(it => it.status !== '已掌握').length;

            const historyList = document.getElementById('historyList');
            if (stats.history.length === 0) {
                historyList.innerHTML = `
                    <div class="empty-state">
                        <div class="empty-icon">📊</div>
                        <p>暂无练习记录</p>
                    </div>
                `;
                return;
            }

            historyList.innerHTML = stats.history.map(h => `
                <div class="history-item ${h.isCorrect ? 'correct' : 'incorrect'}">
                    <div>
                        <strong>${h.isCorrect ? '✓' : '✗'} ${escapeHtml(h.prompt)}</strong>
                        <p style="font-size: 0.9em; color: #666; margin-top: 5px;">
                            ${new Date(h.timestamp).toLocaleString('zh-CN')}
                        </p>
                    </div>
                    <div style="text-align: right;">
                        <p style="font-size: 0.9em;">你的答案：${escapeHtml(h.userAnswer || '')}</p>
                        <p style="font-size: 0.9em; color: ${h.isCorrect ? '#4CAF50' : '#f44336'};">
                            ${h.isCorrect ? '正确' : '正确：' + escapeHtml(h.correctAnswer || '')}
                        </p>
                    </div>
                </div>
            `).join('');
        }

        /* ================================================================
         * 题目管理
         * ================================================================ */
        function restoreBank() {
            if (usingCustomBank) {
                if (confirm('确定要放弃自定义题库、恢复内置 625 题库吗？\n（练习统计与错题本数据不受影响）')) {
                    localStorage.removeItem('toeflBuildQuestions');
                    localStorage.removeItem('toeflBuildBankFlag');
                    usingCustomBank = false;
                    initBank();
                    buildFilterOptions();
                    resetToBankPool();
                    renderQuestionList();
                    showToast('已恢复内置题库');
                }
            } else {
                showToast('当前已在使用内置题库（' + questions.length + ' 题）');
            }
        }

        function renderQuestionList() {
            const container = document.getElementById('questionList');
            const emptyEl = document.getElementById('manageEmptyState');
            const searchTerm = document.getElementById('searchBox').value.toLowerCase();

            if (questions.length === 0) {
                container.innerHTML = '';
                emptyEl.style.display = 'block';
                return;
            }

            emptyEl.style.display = 'none';

            const filtered = questions.filter(q =>
                (q.prompt || '').toLowerCase().includes(searchTerm) ||
                (q.questionText || '').toLowerCase().includes(searchTerm) ||
                (q.correctAnswer || '').toLowerCase().includes(searchTerm)
            );

            container.innerHTML = filtered.slice(0, 200).map(q => {
                const dims = [q.scenario, q.sentenceFunction, q.sentenceStructure, q.purpose].filter(Boolean);
                return `
                <div class="question-list-item" onclick="goToQuestion('${escapeJs(q.id)}')">
                    <strong>题目 ${questions.indexOf(q) + 1} <span style="font-weight:400;color:#bbb;font-size:0.85em;">${escapeHtml(q.id)}</span></strong>
                    <p style="color: #666; margin: 5px 0;">${escapeHtml(q.prompt || '')}</p>
                    <p style="font-size: 0.9em; color: #999;">${escapeHtml(q.questionText || '')}</p>
                    ${dims.length ? `<div class="dim-tags" style="margin-top:6px;">${dims.map(d => `<span class="dim-tag" style="background:#E8F0EC;color:#006633;">${escapeHtml(d)}</span>`).join('')}</div>` : ''}
                </div>
            `;
            }).join('') + (filtered.length > 200 ? `<p style="text-align:center;color:#999;padding:15px;">仅显示前 200 条，共 ${filtered.length} 条匹配，请用关键词缩小范围</p>` : '');
        }

        window.goToQuestion = function (id) {
            const index = questions.findIndex(q => q.id === id);
            if (index !== -1) {
                setSession = null;
                pool = { list: questions.slice(), mode: 'bank', label: '' };
                renderModeBanner();
                updateResultCount(questions.length);
                currentQuestionIndex = index;
                switchTab('practice');
                displayQuestion();
            }
        };

        function openAddModal() {
            document.getElementById('modalTitle').textContent = '添加题目';
            document.getElementById('questionForm').reset();
            document.getElementById('questionModal').classList.add('show');
        }

        function closeModal() {
            document.getElementById('questionModal').classList.remove('show');
        }

        function saveQuestion(e) {
            e.preventDefault();

            const newQuestion = {
                id: 'q' + Date.now(),
                prompt: document.getElementById('formPrompt').value,
                questionText: document.getElementById('formQuestion').value,
                options: document.getElementById('formOptions').value.split(',').map(o => o.trim()).filter(Boolean),
                correctAnswer: document.getElementById('formAnswer').value,
                answerOrder: [],
                source: '手动添加',
                date: wbToday(),
                scenario: '', polarity: '', sentenceFunction: '', sentenceStructure: '', grammar: '', purpose: '',
                explanation: ''
            };

            questions.push(newQuestion);
            persistBank();
            closeModal();
            renderQuestionList();
            showToast('题目添加成功！');
        }

        function exportQuestions() {
            const dataStr = JSON.stringify(questions, null, 2);
            const blob = new Blob([dataStr], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `toefl-build-sentence-questions-${new Date().toISOString().split('T')[0]}.json`;
            a.click();
            showToast('题目已导出');
        }

        /* ================================================================
         * 工具函数
         * ================================================================ */
        function escapeHtml(str) {
            if (str == null) return '';
            return String(str)
                .replace(/&/g, '&amp;')
                .replace(/</g, '&lt;')
                .replace(/>/g, '&gt;')
                .replace(/"/g, '&quot;')
                .replace(/'/g, '&#39;');
        }

        function escapeJs(str) {
            if (str == null) return '';
            return String(str).replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/"/g, '&quot;');
        }

        function showToast(message) {
            const toast = document.getElementById('toast');
            toast.textContent = message;
            toast.classList.add('show');
            clearTimeout(toast._timer);
            toast._timer = setTimeout(() => {
                toast.classList.remove('show');
            }, 3000);
        }
