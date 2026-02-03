// 全局变量
let users = [];
let winners = [];
let currentLottery = null;
let isLotteryRunning = false;
let lotteryInterval = null;
let bossKeepInPool = true; // 老板中奖后是否留在抽奖池
let autoStopEnabled = false; // 是否自动停止
let autoStopDuration = 5; // 自动停止时长（秒）
let autoStopTimer = null; // 自动停止定时器
let rotationInterval = null; // 旋转加速定时器
let axisShuffleTimer = null; // 抽奖时随机轴定时器
let spinEaseRaf = null; // 自转减速动画
let recreateGlobeTimer = null; // 停止后重建球体定时器
let currentSpinAxis = { x: 0, y: 1, z: 0 };
let currentSpinDuration = 30;
const GLOBE_RADIUS = 210;
const FULLSCREEN_GLOBE_RADIUS = 260;
const DEFAULT_WINNER_REVEAL_DURATION = 10;
const DEFAULT_MULTI_REVEAL_DELAY = 1;

// 鼠标拖动旋转相关变量
let isDragging = false;
let previousMousePosition = { x: 0, y: 0 };
let globeRotation = { x: 0, y: 0 };
let autoRotate = true; // 是否自动旋转
let lastDragTime = 0;
let dragVelocity = { x: 0, y: 0 };
let awards = [];
let selectedAwardId = null;
let lastDraw = null;
let currentDrawAwardId = null;
let userSearchTerm = '';
const AWARD_IMAGE_DIR = 'images';
const DEFAULT_AWARD_IMAGE = 'gift.jpg';
let pendingAwardImage = {
    type: 'preset',
    preset: DEFAULT_AWARD_IMAGE,
    data: null
};
let winnerRevealDuration = DEFAULT_WINNER_REVEAL_DURATION;
let winnerMultiDelay = DEFAULT_MULTI_REVEAL_DELAY;

const LANG_STORAGE_KEY = 'lotteryLang';
let currentLang = localStorage.getItem(LANG_STORAGE_KEY) || null;

function t(key, vars = {}) {
    const dict = (window.I18N && window.I18N[currentLang]) || window.I18N?.[window.I18N_DEFAULT] || {};
    let text = dict[key] || key;
    Object.keys(vars).forEach((k) => {
        text = text.replace(new RegExp(`\\{${k}\\}`, 'g'), vars[k]);
    });
    return text;
}

function applyI18n() {
    const defaultLang = window.I18N_DEFAULT || 'zh';
    if (!currentLang || !window.I18N || !window.I18N[currentLang]) {
        const browserLang = (navigator.language || '').toLowerCase();
        currentLang = browserLang.startsWith('en') ? 'en' : defaultLang;
        localStorage.setItem(LANG_STORAGE_KEY, currentLang);
    }

    document.querySelectorAll('[data-i18n]').forEach((el) => {
        el.textContent = t(el.dataset.i18n);
    });
    document.querySelectorAll('[data-i18n-placeholder]').forEach((el) => {
        el.setAttribute('placeholder', t(el.dataset.i18nPlaceholder));
    });
    document.querySelectorAll('[data-i18n-aria]').forEach((el) => {
        el.setAttribute('aria-label', t(el.dataset.i18nAria));
    });

    document.title = t('app_title');
    const languageSelect = $('#languageSelect');
    if (languageSelect.length) {
        languageSelect.val(currentLang);
    }

    const drawerToggle = $('#drawerToggle');
    if (drawerToggle.length) {
        const isOpen = $('body').hasClass('drawer-open');
        drawerToggle.attr('aria-label', isOpen ? t('drawer_close') : t('drawer_open'));
    }

    updateFullscreenState();
    updateLotteryStatusText();
    updateTables();
    renderAwards();
    updateStatsPage();
}

function applyWinnerAnimationSettings() {
    const root = document.documentElement;
    root.style.setProperty('--winner-reveal-duration', `${winnerRevealDuration}s`);
    root.style.setProperty('--winner-multi-duration', `${winnerRevealDuration}s`);
    root.style.setProperty('--winner-multi-delay', `${winnerMultiDelay}s`);
}

function updateLotteryStatusText() {
    if (isLotteryRunning) {
        return;
    }
    if (currentLottery) {
        $('#lotteryNumber').text(currentLottery.number);
        $('#lotteryName').text(currentLottery.name);
        return;
    }
    $('#lotteryNumber').text('000');
    $('#lotteryName').text(t('lottery_ready'));
    $('#winnerPopups').empty();
    $('body').removeClass('multi-winner');
    $('#lotteryName').removeClass('boss-name');
}

// 初始化
$(document).ready(function() {
    loadData();
    updateTables();
    updateStats();
    createParticles();
    syncSettingsUI();
    createGlobe(); // 初始化球体
    initMouseDrag(); // 初始化鼠标拖动

    // 事件监听
    $('#startBtn').click(startLottery);
    $('#stopBtn').click(stopLottery);
    $('#resetBtn').click(resetLottery);
    $('#undoBtn').click(undoLastDraw);
    $('#addForm').submit(addUser);
    $('#userExcel').change(importUsers);
    $('#bossExcel').change(importBosses);
    $('#saveEditBtn').click(saveEdit);
    $('#clearDataBtn').click(clearData);
    $('#awardForm').submit(addAward);
    $('#userSearchInput').on('input', function() {
        userSearchTerm = $(this).val().trim().toLowerCase();
        updateUserTable();
    });
    $('#awardImagePreset').on('change', function() {
        pendingAwardImage = {
            type: 'preset',
            preset: $(this).val() || DEFAULT_AWARD_IMAGE,
            data: null
        };
    });
    $('#awardImageUpload').on('change', function() {
        const file = this.files && this.files[0];
        if (!file) {
            pendingAwardImage = {
                type: 'preset',
                preset: $('#awardImagePreset').val() || DEFAULT_AWARD_IMAGE,
                data: null
            };
            return;
        }
        const reader = new FileReader();
        reader.onload = () => {
            pendingAwardImage = {
                type: 'custom',
                preset: $('#awardImagePreset').val() || DEFAULT_AWARD_IMAGE,
                data: reader.result
            };
        };
        reader.readAsDataURL(file);
    });
    $('#awardList').on('click', '.award-card', function() {
        if (isLotteryRunning) return;
        if ($(this).hasClass('disabled')) return;
        selectAward(Number($(this).data('id')));
    });
    $('#awardManageList').on('click', '[data-action="award-delete"]', function() {
        const awardId = Number($(this).closest('.award-manage-item').data('id'));
        deleteAward(awardId);
    });
    $('#awardManageList').on('change', 'select[data-action="award-preset"]', function() {
        const awardId = Number($(this).closest('.award-manage-item').data('id'));
        const award = awards.find((item) => item.id === awardId);
        if (!award) return;
        award.imageType = 'preset';
        award.imagePreset = $(this).val() || DEFAULT_AWARD_IMAGE;
        award.imageData = null;
        renderAwards();
        updateStatsPage();
    });
    $('#awardManageList').on('change', 'input[data-action="award-upload"]', function() {
        const file = this.files && this.files[0];
        if (!file) return;
        const awardId = Number($(this).closest('.award-manage-item').data('id'));
        const award = awards.find((item) => item.id === awardId);
        if (!award) return;
        const reader = new FileReader();
        reader.onload = () => {
            award.imageType = 'custom';
            award.imageData = reader.result;
            renderAwards();
            updateStatsPage();
        };
        reader.readAsDataURL(file);
    });
    $('#awardManageList').on('click', '[data-action="award-inc"]', function() {
        const item = $(this).closest('.award-manage-item');
        adjustAwardValue(Number(item.data('id')), $(this).data('field'), 1);
    });
    $('#awardManageList').on('click', '[data-action="award-dec"]', function() {
        const item = $(this).closest('.award-manage-item');
        adjustAwardValue(Number(item.data('id')), $(this).data('field'), -1);
    });
    $('#awardManageList').on('change', 'input[data-field]', function() {
        const item = $(this).closest('.award-manage-item');
        updateAwardValue(Number(item.data('id')), $(this).data('field'), $(this).val());
    });
    $('#awardPanel').on('mouseenter', function() {
        $(this).removeClass('show-selected');
    });
    $('#awardList').on('mouseleave', function() {
        if (selectedAwardId) {
            $('#awardPanel').addClass('show-selected');
        }
    });
    $('#bossKeepInPool').change(function() {
        bossKeepInPool = $(this).is(':checked');
        localStorage.setItem('bossKeepInPool', bossKeepInPool);
        showAlert(t('settings_saved'), 'success');
    });
    $('#autoStopEnabled').change(function() {
        autoStopEnabled = $(this).is(':checked');
        localStorage.setItem('autoStopEnabled', autoStopEnabled);
        $('#autoStopDurationContainer').toggle(autoStopEnabled);
        showAlert(t('settings_saved'), 'success');
    });
    $('#autoStopDuration').change(function() {
        autoStopDuration = parseInt($(this).val()) || 5;
        localStorage.setItem('autoStopDuration', autoStopDuration);
    });
    $('#winnerRevealDuration').change(function() {
        const value = parseFloat($(this).val());
        winnerRevealDuration = Number.isFinite(value) && value > 0 ? value : DEFAULT_WINNER_REVEAL_DURATION;
        localStorage.setItem('winnerRevealDuration', winnerRevealDuration);
        applyWinnerAnimationSettings();
        showAlert(t('settings_saved'), 'success');
    });
    $('#winnerMultiDelay').change(function() {
        const value = parseFloat($(this).val());
        winnerMultiDelay = Number.isFinite(value) && value >= 0 ? value : DEFAULT_MULTI_REVEAL_DELAY;
        localStorage.setItem('winnerMultiDelay', winnerMultiDelay);
        applyWinnerAnimationSettings();
        showAlert(t('settings_saved'), 'success');
    });

    $('#drawerToggle').on('click', function() {
        $('body').toggleClass('drawer-open');
        const isOpen = $('body').hasClass('drawer-open');
        $(this).attr('aria-label', isOpen ? t('drawer_close') : t('drawer_open'));
    });
    $('#drawerClose').on('click', function() {
        $('body').removeClass('drawer-open');
    });
    $('#drawerBackdrop').on('click', function() {
        $('body').removeClass('drawer-open');
    });

    $('button[data-bs-toggle="tab"]').on('shown.bs.tab', function(e) {
        const target = $(e.target).attr('data-bs-target');
        if (target === '#stats') {
            updateStatsPage();
        } else if (target === '#settings') {
            updateStats();
            syncSettingsUI();
        } else if (target === '#lottery') {
            updateResultTable();
        } else if (target === '#manage') {
            updateUserTable();
        }
    });

    $('#fullscreenToggle').on('click', function() {
        if (!document.fullscreenElement) {
            document.documentElement.requestFullscreen().catch(() => {});
        } else {
            document.exitFullscreen();
        }
    });

    document.addEventListener('fullscreenchange', updateFullscreenState);
    updateFullscreenState();

    $('#languageSelect').on('change', function() {
        currentLang = $(this).val();
        localStorage.setItem(LANG_STORAGE_KEY, currentLang);
        applyI18n();
    });

    applyI18n();
    applyWinnerAnimationSettings();
});

// 加载数据
function loadData() {
    users = JSON.parse(localStorage.getItem('lotteryUsers')) || [];
    winners = JSON.parse(localStorage.getItem('lotteryWinners')) || [];
    bossKeepInPool = localStorage.getItem('bossKeepInPool') !== 'false'; // 默认为true
    autoStopEnabled = localStorage.getItem('autoStopEnabled') === 'true';
    autoStopDuration = parseInt(localStorage.getItem('autoStopDuration')) || 5;
    const savedRevealDuration = parseFloat(localStorage.getItem('winnerRevealDuration'));
    const savedMultiDelay = parseFloat(localStorage.getItem('winnerMultiDelay'));
    winnerRevealDuration = Number.isFinite(savedRevealDuration) && savedRevealDuration > 0
        ? savedRevealDuration
        : DEFAULT_WINNER_REVEAL_DURATION;
    winnerMultiDelay = Number.isFinite(savedMultiDelay) && savedMultiDelay >= 0
        ? savedMultiDelay
        : DEFAULT_MULTI_REVEAL_DELAY;
    awards = (JSON.parse(localStorage.getItem('lotteryAwards')) || []).map(normalizeAward);
    selectedAwardId = localStorage.getItem('selectedAwardId') || null;
}

// 保存数据
function saveData() {
    localStorage.setItem('lotteryUsers', JSON.stringify(users));
    localStorage.setItem('lotteryWinners', JSON.stringify(winners));
    localStorage.setItem('lotteryAwards', JSON.stringify(awards));
    if (selectedAwardId) {
        localStorage.setItem('selectedAwardId', selectedAwardId);
    } else {
        localStorage.removeItem('selectedAwardId');
    }
}

// 更新表格
function updateTables() {
    updateUserTable();
    updateResultTable();
    renderAwards();
    updateStatsPage();
    $('#undoBtn').prop('disabled', !lastDraw);
}

// 更新人员表格
function updateUserTable() {
    const tbody = $('#userTable');
    tbody.empty();

    const filteredUsers = userSearchTerm
        ? users.filter((user) => {
            const name = (user.name || '').toLowerCase();
            const number = (user.number || '').toLowerCase();
            return name.includes(userSearchTerm) || number.includes(userSearchTerm);
        })
        : users;

    filteredUsers.forEach((user, index) => {
        const tr = $('<tr></tr>');
        tr.append($('<td></td>').text(user.name));
        tr.append($('<td></td>').text(user.number));

        const typeBadge = user.type === 'boss'
            ? $(`<span class="badge badge-boss"><i class="fa fa-star"></i> ${t('type_boss')}</span>`)
            : $(`<span class="badge bg-secondary">${t('type_user')}</span>`);
        tr.append($('<td></td>').append(typeBadge));

        const actions = $('<td></td>');
        const originalIndex = users.indexOf(user);
        actions.append($('<button class="btn btn-sm btn-primary me-1" onclick="editUser(' + originalIndex + ')"><i class="fa fa-pencil"></i></button>'));
        actions.append($('<button class="btn btn-sm btn-danger" onclick="deleteUser(' + originalIndex + ')"><i class="fa fa-trash"></i></button>'));
        tr.append(actions);

        tbody.append(tr);
    });
}

// 更新结果表格
function updateResultTable() {
    const renderTable = (tbody) => {
        tbody.empty();
        winners.forEach((winner, index) => {
            const tr = $('<tr></tr>');
            tr.append($('<td></td>').text(index + 1));
            const nameCell = $('<td></td>').append($('<span></span>').text(winner.name));
            if (winner.type === 'boss' && bossKeepInPool) {
                nameCell.append(` <span class="badge badge-boss-win">${t('boss_win_note')}</span>`);
            }
            tr.append(nameCell);
            tr.append($('<td></td>').text(winner.number));

            const typeBadge = winner.type === 'boss'
                ? $(`<span class="badge badge-boss"><i class="fa fa-star"></i> ${t('type_boss')}</span>`)
                : $(`<span class="badge bg-secondary">${t('type_user')}</span>`);
            tr.append($('<td></td>').append(typeBadge));
            tr.append($('<td></td>').text(winner.awardName || '-'));
            tr.append($('<td></td>').text(winner.awardPrize || '-'));

            // 添加中奖时间
            tr.append($('<td></td>').text(winner.winTime || '-'));

            tbody.append(tr);
        });
    };

    const drawerBody = $('#resultTable');
    if (drawerBody.length) {
        renderTable(drawerBody);
    }

    const inlineBody = $('#resultTableInline');
    if (inlineBody.length) {
        renderTable(inlineBody);
    }

    $('#resultCount').text(winners.length);
    $('#resultCountInline').text(winners.length);
}

// 更新统计信息
function updateStats() {
    const remainingUsers = users.filter(user => {
        const isWinner = winners.some(winner => winner.id === user.id);
        // 如果开启老板留池，老板即使中奖也不移除
        if (bossKeepInPool && user.type === 'boss') {
            return true;
        }
        return !isWinner;
    });

    $('#remainingCount').text(remainingUsers.length);
    $('#winnerCount').text(winners.length);
    $('#bossCount').text(users.filter(u => u.type === 'boss').length);
    $('#userCount').text(users.filter(u => u.type === 'user').length);
    updateStatsPage();
}

function updateStatsPage() {
    const statsTotalWinners = $('#statsTotalWinners');
    if (!statsTotalWinners.length) return;

    const totalWinners = winners.length;
    const bossWins = winners.filter(w => w.type === 'boss').length;
    statsTotalWinners.text(totalWinners);
    $('#statsTotalAwards').text(awards.length);
    $('#statsBossWins').text(bossWins);

    const awardTable = $('#statsAwardTable');
    awardTable.empty();
    if (awards.length === 0) {
        awardTable.append(`<tr><td colspan="4" class="text-muted">${t('award_empty')}</td></tr>`);
    } else {
        awards.forEach((award) => {
            const drawn = winners.filter(w => w.awardId === award.id).length;
            const remaining = Number.isFinite(award.remaining) ? award.remaining : Math.max(0, award.total - drawn);
            const tr = $('<tr></tr>');
            tr.append($('<td></td>').text(award.name));
            tr.append($('<td></td>').text(award.prize));
            tr.append($('<td></td>').text(drawn));
            tr.append($('<td></td>').text(remaining));
            awardTable.append(tr);
        });
    }

    const winnerTable = $('#statsWinnerTable');
    winnerTable.empty();
    winners.forEach((winner, index) => {
        const tr = $('<tr></tr>');
        tr.append($('<td></td>').text(index + 1));
        const nameCell = $('<td></td>').append($('<span></span>').text(winner.name));
        if (winner.type === 'boss' && bossKeepInPool) {
            nameCell.append(` <span class="badge badge-boss-win">${t('boss_win_note')}</span>`);
        }
        tr.append(nameCell);
        tr.append($('<td></td>').text(winner.number));
        const typeLabel = winner.type === 'boss' ? t('type_boss') : t('type_user');
        tr.append($('<td></td>').text(typeLabel));
        tr.append($('<td></td>').text(winner.awardName || '-'));
        tr.append($('<td></td>').text(winner.awardPrize || '-'));
        tr.append($('<td></td>').text(winner.winTime || '-'));
        winnerTable.append(tr);
    });
}

// 更新名字墙 - 已移除，改用3D球体显示
function updateNameWall() {
    // 重新创建球体以更新人员列表
    if (!isLotteryRunning) {
        createGlobe();
    }
}

function getRandomAxis() {
    let x = Math.random() * 2 - 1;
    let y = Math.random() * 2 - 1;
    let z = Math.random() * 2 - 1;
    const length = Math.hypot(x, y, z) || 1;
    return {
        x: (x / length).toFixed(3),
        y: (y / length).toFixed(3),
        z: (z / length).toFixed(3)
    };
}

function setGlobeSpin(globe, { axis = { x: 0, y: 1, z: 0 }, duration = 30, delay = 0 } = {}) {
    currentSpinAxis = { x: axis.x, y: axis.y, z: axis.z };
    currentSpinDuration = duration;
    globe.css({
        '--axis-x': axis.x,
        '--axis-y': axis.y,
        '--axis-z': axis.z,
        '--spin-duration': `${duration}s`,
        '--spin-delay': `${delay}s`
    });
}

function updateDragCounterRotation(globe) {
    globe.css({
        '--drag-x': `${globeRotation.x}deg`,
        '--drag-y': `${globeRotation.y}deg`
    });
}

function normalizeAward(award) {
    const total = Number(award.total) || 0;
    const perDraw = Number(award.perDraw) || 1;
    let remaining = Number(award.remaining);
    if (!Number.isFinite(remaining)) {
        remaining = total;
    }
    const imageType = award.imageType === 'custom' && award.imageData ? 'custom' : 'preset';
    const imagePreset = award.imagePreset || DEFAULT_AWARD_IMAGE;
    const imageData = award.imageData || null;
    return {
        ...award,
        total,
        perDraw,
        remaining: Math.max(0, remaining),
        imageType,
        imagePreset,
        imageData
    };
}

function getSelectedAward() {
    if (!selectedAwardId) return null;
    return awards.find((award) => award.id === selectedAwardId) || null;
}

function getAwardImageSrc(award) {
    if (!award) return `${AWARD_IMAGE_DIR}/${DEFAULT_AWARD_IMAGE}`;
    if (award.imageType === 'custom' && award.imageData) {
        return award.imageData;
    }
    const preset = award.imagePreset || DEFAULT_AWARD_IMAGE;
    return `${AWARD_IMAGE_DIR}/${preset}`;
}

function updateSelectedAwardOverlay() {
    const award = getSelectedAward();
    if (!award) {
        $('#selectedAwardName').text('-');
        $('#selectedAwardPrize').text('-');
        $('#selectedAwardPerDraw').text('0');
        $('#selectedAwardRemaining').text('0');
        $('#selectedAwardImage').attr('src', `${AWARD_IMAGE_DIR}/${DEFAULT_AWARD_IMAGE}`);
        $('#currentAwardImage').attr('src', `${AWARD_IMAGE_DIR}/${DEFAULT_AWARD_IMAGE}`);
        return;
    }
    $('#selectedAwardName').text(award.name);
    $('#selectedAwardPrize').text(award.prize);
    $('#selectedAwardPerDraw').text(award.perDraw);
    $('#selectedAwardRemaining').text(award.remaining);
    const imageSrc = getAwardImageSrc(award);
    $('#selectedAwardImage').attr('src', imageSrc);
    $('#currentAwardImage').attr('src', imageSrc);
}

function renderAwards() {
    awards = awards.map(normalizeAward);
    if (selectedAwardId && !awards.find((award) => award.id === selectedAwardId)) {
        selectedAwardId = null;
    }
    const list = $('#awardList');
    const manageList = $('#awardManageList');
    if (list.length) {
        list.empty();
        if (awards.length === 0) {
            list.append(`<div class="text-muted">${t('award_empty')}</div>`);
        } else {
            awards.forEach((award) => {
                const isSelected = award.id === selectedAwardId;
                const isDisabled = award.remaining <= 0 || award.perDraw <= 0 || award.remaining < award.perDraw;
                const imageSrc = getAwardImageSrc(award);
                const card = $(`
                    <div class="award-card ${isSelected ? 'selected' : ''} ${isDisabled ? 'disabled' : ''}" data-id="${award.id}">
                        <div class="award-badge">
                            <img src="${imageSrc}" alt="${award.name}">
                        </div>
                        <div class="award-info">
                            <div class="award-title">${award.name}</div>
                            <div class="award-prize">${award.prize}</div>
                            <div class="award-meta">
                                <span>${t('award_per_draw')} ${award.perDraw}</span>
                                <span class="award-remaining">${t('award_remaining')} ${award.remaining}</span>
                            </div>
                        </div>
                    </div>
                `);
                list.append(card);
            });
        }
    }

    if (manageList.length) {
        manageList.empty();
        if (awards.length === 0) {
            manageList.append(`<div class="text-muted">${t('award_empty')}</div>`);
        } else {
            awards.forEach((award) => {
                const imageSrc = getAwardImageSrc(award);
                const presetValue = award.imagePreset || DEFAULT_AWARD_IMAGE;
                const item = $(`
                    <div class="award-manage-item" data-id="${award.id}">
                        <div class="award-manage-header">
                            <div class="award-manage-info">
                                <div class="award-manage-thumb">
                                    <img src="${imageSrc}" alt="${award.name}">
                                </div>
                                <div class="award-manage-body">
                                    <div class="award-manage-line">
                                        <span class="award-manage-title">${award.name}</span>
                                        <span class="award-manage-divider">•</span>
                                        <span class="award-manage-prize">${award.prize}</span>
                                    </div>
                                    <div class="award-manage-upload">
                                        <div class="award-manage-upload-item">
                                            <label class="form-label mb-1">${t('award_image_preset')}</label>
                                            <select class="form-select form-select-sm" data-action="award-preset">
                                                <option value="gift.jpg">${t('award_image_gift')}</option>
                                                <option value="money.jpg">${t('award_image_money')}</option>
                                                <option value="phone.jpg">${t('award_image_phone')}</option>
                                            </select>
                                        </div>
                                        <div class="award-manage-upload-item">
                                            <label class="form-label mb-1">${t('award_image_upload')}</label>
                                            <input type="file" class="form-control form-control-sm" data-action="award-upload" accept="image/*">
                                        </div>
                                    </div>
                                    <div class="award-manage-controls">
                                        <div class="award-manage-control-labels">
                                            <span>${t('award_total')}</span>
                                            <span>${t('award_per_draw')}</span>
                                        </div>
                                        <div class="award-manage-control-inputs">
                                            <div class="award-manage-control">
                                                <button class="btn btn-sm btn-outline-secondary" data-action="award-dec" data-field="total">-</button>
                                                <input type="number" class="form-control form-control-sm" data-field="total" value="${award.total}" min="1">
                                                <button class="btn btn-sm btn-outline-secondary" data-action="award-inc" data-field="total">+</button>
                                            </div>
                                            <div class="award-manage-control">
                                                <button class="btn btn-sm btn-outline-secondary" data-action="award-dec" data-field="perDraw">-</button>
                                                <input type="number" class="form-control form-control-sm" data-field="perDraw" value="${award.perDraw}" min="1">
                                                <button class="btn btn-sm btn-outline-secondary" data-action="award-inc" data-field="perDraw">+</button>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="award-manage-actions">
                                <button class="btn btn-sm btn-outline-danger" data-action="award-delete"><i class="fa fa-trash"></i></button>
                            </div>
                        </div>
                    </div>
                `);
                item.find('select[data-action="award-preset"]').val(presetValue);
                manageList.append(item);
            });
        }
    }

    updateSelectedAwardOverlay();
    saveData();
}

function selectAward(awardId) {
    selectedAwardId = awardId;
    renderAwards();
}

function addAward(e) {
    e.preventDefault();
    const name = $('#awardName').val().trim();
    const prize = $('#awardPrize').val().trim();
    const total = parseInt($('#awardTotal').val(), 10);
    const perDraw = parseInt($('#awardPerDraw').val(), 10);
    if (!name || !prize || !Number.isFinite(total) || !Number.isFinite(perDraw) || total <= 0 || perDraw <= 0) {
        showAlert(t('award_required'), 'warning');
        return;
    }
    if (perDraw > total) {
        showAlert(t('award_per_draw_exceed_total'), 'warning');
        return;
    }
    const imageType = pendingAwardImage.type === 'custom' && pendingAwardImage.data ? 'custom' : 'preset';
    const imagePreset = pendingAwardImage.preset || DEFAULT_AWARD_IMAGE;
    const imageData = pendingAwardImage.data || null;
    const award = normalizeAward({
        id: Date.now() + Math.random(),
        name,
        prize,
        total,
        perDraw,
        remaining: total,
        imageType,
        imagePreset,
        imageData
    });
    awards.push(award);
    if (!selectedAwardId) {
        selectedAwardId = award.id;
    }
    $('#awardForm')[0].reset();
    $('#awardImageUpload').val('');
    pendingAwardImage = {
        type: 'preset',
        preset: DEFAULT_AWARD_IMAGE,
        data: null
    };
    renderAwards();
    updateStatsPage();
}

function updateAwardValue(awardId, field, value) {
    const award = awards.find((item) => item.id === awardId);
    if (!award) return;
    const numericValue = parseInt(value, 10);
    if (!Number.isFinite(numericValue) || numericValue <= 0) return;
    if (field === 'total') {
        const awarded = award.total - award.remaining;
        if (numericValue < awarded) {
            showAlert(t('award_total_too_small'), 'warning');
            return;
        }
        if (numericValue < award.perDraw) {
            showAlert(t('award_per_draw_exceed_total'), 'warning');
            return;
        }
        award.total = numericValue;
        award.remaining = numericValue - awarded;
    } else if (field === 'perDraw') {
        if (numericValue > award.total) {
            showAlert(t('award_per_draw_exceed_total'), 'warning');
            return;
        }
        award.perDraw = numericValue;
    }
    renderAwards();
    updateStatsPage();
}

function adjustAwardValue(awardId, field, delta) {
    const award = awards.find((item) => item.id === awardId);
    if (!award) return;
    const currentValue = field === 'total' ? award.total : award.perDraw;
    const nextValue = Math.max(1, currentValue + delta);
    updateAwardValue(awardId, field, nextValue);
}

function deleteAward(awardId) {
    const award = awards.find((item) => item.id === awardId);
    if (!award) return;
    if (!confirm(t('award_delete_confirm'))) {
        return;
    }
    awards = awards.filter((item) => item.id !== awardId);
    if (selectedAwardId === awardId) {
        selectedAwardId = null;
    }
    renderAwards();
    updateStatsPage();
}

function getGlobeRadius() {
    return document.fullscreenElement ? FULLSCREEN_GLOBE_RADIUS : GLOBE_RADIUS;
}

function updateFullscreenState() {
    const isFullscreen = !!document.fullscreenElement;
    $('body').toggleClass('fullscreen-open', isFullscreen);
    $('html').toggleClass('fullscreen-open', isFullscreen);
    const button = $('#fullscreenToggle');
    if (button.length) {
        const icon = isFullscreen ? 'compress' : 'arrows-alt';
        const text = isFullscreen ? t('fullscreen_exit') : t('fullscreen_enter');
        button.html(`<i class="fa fa-${icon}" aria-hidden="true"></i> ${text}`);
    }
    if (!isLotteryRunning) {
        createGlobe();
    }
}

function getAxisFromVelocity(vx, vy) {
    const ax = -vy;
    const ay = vx;
    const length = Math.hypot(ax, ay) || 1;
    return {
        x: (ax / length).toFixed(3),
        y: (ay / length).toFixed(3),
        z: 0
    };
}

function stopSpinEase() {
    if (spinEaseRaf) {
        cancelAnimationFrame(spinEaseRaf);
        spinEaseRaf = null;
    }
}

function easeSpinDuration(globe, fromDuration, toDuration, durationMs = 1600) {
    stopSpinEase();
    const start = performance.now();
    const delta = toDuration - fromDuration;

    const tick = (now) => {
        const t = Math.min(1, (now - start) / durationMs);
        const eased = 1 - Math.pow(1 - t, 3);
        const current = fromDuration + delta * eased;
        currentSpinDuration = current;
        globe.css('--spin-duration', current + 's');
        if (t < 1) {
            spinEaseRaf = requestAnimationFrame(tick);
        } else {
            spinEaseRaf = null;
        }
    };

    spinEaseRaf = requestAnimationFrame(tick);
}

function stopAxisShuffle() {
    if (axisShuffleTimer) {
        clearTimeout(axisShuffleTimer);
        axisShuffleTimer = null;
    }
}

function startAxisShuffle(globe) {
    stopAxisShuffle();
    const tick = () => {
        const axis = getRandomAxis();
        currentSpinAxis = { x: axis.x, y: axis.y, z: axis.z };
        globe.css({
            '--axis-x': axis.x,
            '--axis-y': axis.y,
            '--axis-z': axis.z
        });
        axisShuffleTimer = setTimeout(tick, currentSpinDuration * 1000);
    };
    axisShuffleTimer = setTimeout(tick, currentSpinDuration * 1000);
}

function resumeAutoRotateFromDrag(globe) {
    const speed = Math.hypot(dragVelocity.x, dragVelocity.y);
    const degPerSec = speed * 1000;
    const maxDuration = 30;
    const minDuration = 1.6;
    let duration = maxDuration;
    if (degPerSec > 5) {
        duration = Math.max(minDuration, Math.min(maxDuration, 360 / degPerSec));
    }

    const axis = speed > 0.01 ? getAxisFromVelocity(dragVelocity.x, dragVelocity.y) : currentSpinAxis;
    setGlobeSpin(globe, { axis, duration, delay: -Math.random() * duration });
    globe.addClass('auto-rotate');
    if (duration < maxDuration) {
        easeSpinDuration(globe, duration, maxDuration, 1800);
    }
}

function pickRandomUsers(list, count) {
    const pool = [...list];
    const picked = [];
    const max = Math.min(count, pool.length);
    for (let i = 0; i < max; i++) {
        const index = Math.floor(Math.random() * pool.length);
        picked.push(pool[index]);
        pool.splice(index, 1);
    }
    return picked;
}

// 添加用户
function addUser(e) {
    e.preventDefault();

    const name = $('#addName').val();
    const number = $('#addNumber').val();
    const type = $('#addType').val();

    // 检查号码是否重复
    if (users.some(user => user.number === number)) {
        alert(t('number_exists'));
        return;
    }

    const user = {
        id: Date.now(),
        name: name,
        number: number,
        type: type
    };

    users.push(user);
    saveData();
    updateTables();
    updateStats();
    updateNameWall();

    $('#addForm')[0].reset();
    showAlert(t('add_success'), 'success');
}

// 编辑用户
function editUser(index) {
    const user = users[index];
    $('#editId').val(index);
    $('#editName').val(user.name);
    $('#editNumber').val(user.number);
    $('#editType').val(user.type);
    $('#editModal').modal('show');
}

// 保存编辑
function saveEdit() {
    const index = $('#editId').val();
    const name = $('#editName').val();
    const number = $('#editNumber').val();
    const type = $('#editType').val();

    // 检查号码是否重复（排除自身）
    if (users.some((user, i) => i !== parseInt(index) && user.number === number)) {
        alert(t('number_exists'));
        return;
    }

    users[index] = {
        ...users[index],
        name: name,
        number: number,
        type: type
    };

    saveData();
    updateTables();
    updateStats();
    updateNameWall();
    $('#editModal').modal('hide');
    showAlert(t('update_success'), 'success');
}

// 删除用户
function deleteUser(index) {
    if (confirm(t('delete_confirm'))) {
        users.splice(index, 1);
        saveData();
        updateTables();
        updateStats();
        updateNameWall();
        showAlert(t('delete_success'), 'success');
    }
}

// 导入员工Excel
function importUsers(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        importUsersFromJson(jsonData, 'user');

        // 重置文件输入，允许重复选择同一文件
        $('#userExcel').val('');
    };
    reader.readAsArrayBuffer(file);
}

// 导入老板Excel
function importBosses(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet);

        importUsersFromJson(jsonData, 'boss');

        // 重置文件输入，允许重复选择同一文件
        $('#bossExcel').val('');
    };
    reader.readAsArrayBuffer(file);
}

// 从JSON导入用户
function importUsersFromJson(jsonData, type) {
    let importedCount = 0;
    let duplicateCount = 0;
    let skippedEmpty = 0;

    if (jsonData.length === 0) {
        showAlert(t('excel_empty'), 'warning');
        return;
    }

    // 获取第一行的列名
    const firstRow = jsonData[0];
    const keys = Object.keys(firstRow);

    // 尝试匹配列名，如果匹配不上则使用第一列和第二列
    let nameKey, numberKey;

    // 查找姓名列（增加更多匹配规则）
    nameKey = keys.find(key => {
        const keyLower = key.toLowerCase().trim();
        return keyLower.includes('姓名') ||
               keyLower.includes('name') ||
               keyLower === '名' ||
               keyLower.includes('人员') ||
               keyLower.includes('员工') ||
               keyLower === '名字' ||
               keyLower === '使用者' ||
               keyLower === '职工';
    });

    // 查找号码列（增加更多匹配规则）
    numberKey = keys.find(key => {
        const keyLower = key.toLowerCase().trim();
        return keyLower.includes('号码') ||
               keyLower.includes('number') ||
               keyLower === '号' ||
               keyLower.includes('工号') ||
               keyLower.includes('编号') ||
               keyLower.includes('员工号') ||
               keyLower === 'no' ||
               keyLower === 'id';
    });

    // 如果找不到匹配的列名，尝试通过数据内容智能判断
    if (!nameKey || !numberKey) {
        console.log('列名匹配失败，尝试智能判断...');

        // 遍历所有列，统计每列的数据特征
        const columnStats = keys.map(key => {
            const values = jsonData.map(row => String(row[key] || '')).filter(v => v.trim());
            const numericCount = values.filter(v => /^\d+$/.test(v.trim())).length;
            const chineseCount = values.filter(v => /[\u4e00-\u9fa5]/.test(v)).length;

            return {
                key,
                total: values.length,
                numericRatio: values.length > 0 ? numericCount / values.length : 0,
                chineseRatio: values.length > 0 ? chineseCount / values.length : 0,
                values: values.slice(0, 5) // 取前5个值作为样本
            };
        });

        console.log('列统计信息:', columnStats);

        // 数字比例高的列作为号码列，中文比例高的列作为姓名列
        const numericColumn = columnStats.find(col => col.numericRatio > 0.5);
        const chineseColumn = columnStats.find(col => col.chineseRatio > 0.5);

        if (numericColumn && !numberKey) {
            numberKey = numericColumn.key;
            console.log('通过数字特征识别号码列:', numberKey);
        }
        if (chineseColumn && !nameKey) {
            nameKey = chineseColumn.key;
            console.log('通过中文特征识别姓名列:', nameKey);
        }

        // 如果还是找不到，使用默认规则
        if (!nameKey && keys.length > 0) {
            nameKey = keys[0];
            console.log('使用第一列作为姓名列:', nameKey);
        }
        if (!numberKey && keys.length > 1) {
            numberKey = keys[1];
            console.log('使用第二列作为号码列:', numberKey);
        }
    }

    console.log('最终识别的列:', { nameKey, numberKey, allKeys: keys });

    // 验证识别结果
    if (!nameKey || !numberKey) {
        showAlert(t('excel_columns_missing'), 'danger');
        return;
    }

    if (nameKey === numberKey) {
        showAlert(t('excel_same_column'), 'danger');
        return;
    }

    jsonData.forEach(row => {
        const name = row[nameKey] || '';
        const number = row[numberKey] || '';

        if (name && number) {
            // 检查号码是否重复
            if (!users.some(user => user.number === number)) {
                users.push({
                    id: Date.now() + Math.random(),
                    name: String(name).trim(),
                    number: String(number).trim(),
                    type: type
                });
                importedCount++;
            } else {
                duplicateCount++;
            }
        } else {
            skippedEmpty++;
        }
    });

    // 如果全部都是空数据，给出提示
    if (importedCount === 0 && duplicateCount === 0) {
        showAlert(t('excel_no_valid'), 'warning');
        return;
    }

    saveData();
    updateTables();
    updateStats();
    updateNameWall();

    let message = t('import_success', { count: importedCount });
    if (duplicateCount > 0) {
        message += t('import_skip_duplicate', { count: duplicateCount });
    }
    if (skippedEmpty > 0) {
        message += t('import_skip_empty', { count: skippedEmpty });
    }
    showAlert(message, 'success');
}

// 创建3D球体
function createGlobe() {
    const globeContainer = $('#globeContainer');
    globeContainer.empty();

    const globe = $('<div class="globe"></div>');
    if (autoRotate && !isLotteryRunning) {
        globe.addClass('auto-rotate');
        setGlobeSpin(globe, { axis: { x: 0, y: 1, z: 0 }, duration: 30, delay: -Math.random() * 30 });
    }
    globeContainer.append(globe);
    globeRotation = { x: 0, y: 0 };
    updateDragCounterRotation(globe);

    // 获取所有未中奖的人员
    const remainingUsers = users.filter(user => {
        const isWinner = winners.some(winner => winner.id === user.id);
        if (bossKeepInPool && user.type === 'boss') {
            return true;
        }
        return !isWinner;
    });

    // 在球体上放置名字
    const radius = getGlobeRadius();
    remainingUsers.forEach((user, index) => {
        const nameWrap = $('<div class="globe-name-wrap"></div>');
        const name = $('<div class="globe-name"></div>');
        name.text(user.name);

        if (user.type === 'boss') {
            name.addClass('boss-name');
        }

        // 使用斐波那契球面分布算法均匀分布名字
        const phi = Math.acos(-1 + (2 * index + 1) / remainingUsers.length);
        const theta = Math.sqrt(remainingUsers.length * Math.PI) * phi;

        const x = radius * Math.cos(theta) * Math.sin(phi);
        const y = radius * Math.sin(theta) * Math.sin(phi);
        const z = radius * Math.cos(phi);

        nameWrap.css({
            'transform': `translate3d(${x}px, ${y}px, ${z}px)`
        });

        nameWrap.append(name);
        globe.append(nameWrap);
    });
}

// 初始化鼠标拖动功能
function initMouseDrag() {
    const globeContainer = $('#globeContainer');

    // 鼠标事件
    globeContainer.on('mousedown', '.globe', function(e) {
        if (isLotteryRunning) return; // 抽奖中禁止拖动

        const globe = $(this);
        isDragging = true;
        autoRotate = false; // 停止自动旋转
        stopSpinEase();
        globe.removeClass('auto-rotate rotating accelerating');

        previousMousePosition = {
            x: e.clientX,
            y: e.clientY
        };
        lastDragTime = performance.now();
        dragVelocity = { x: 0, y: 0 };

        $(document).on('mousemove.drag', function(e) {
            if (!isDragging) return;

            const deltaX = e.clientX - previousMousePosition.x;
            const deltaY = e.clientY - previousMousePosition.y;
            const now = performance.now();
            const dt = Math.max(8, now - lastDragTime);

            globeRotation.y += deltaX * 0.5;
            globeRotation.x -= deltaY * 0.5;

            // 限制X轴旋转角度，避免翻转
            globeRotation.x = Math.max(-60, Math.min(60, globeRotation.x));
            updateDragCounterRotation(globe);

            previousMousePosition = {
                x: e.clientX,
                y: e.clientY
            };
            const vx = (deltaX * 0.5) / dt;
            const vy = (-deltaY * 0.5) / dt;
            dragVelocity.x = dragVelocity.x * 0.8 + vx * 0.2;
            dragVelocity.y = dragVelocity.y * 0.8 + vy * 0.2;
            lastDragTime = now;
        });

        $(document).on('mouseup.drag', function() {
            isDragging = false;
            $(document).off('mousemove.drag mouseup.drag');
            if (!isLotteryRunning) {
                autoRotate = true;
                resumeAutoRotateFromDrag(globe);
            }
        });
    });

    // 触摸设备支持
    globeContainer.on('touchstart', '.globe', function(e) {
        if (isLotteryRunning) return;

        const globe = $(this);
        isDragging = true;
        autoRotate = false;
        stopSpinEase();
        globe.removeClass('auto-rotate rotating accelerating');

        const touch = e.originalEvent.touches[0];
        previousMousePosition = {
            x: touch.clientX,
            y: touch.clientY
        };
        lastDragTime = performance.now();
        dragVelocity = { x: 0, y: 0 };

        $(document).on('touchmove.drag', function(e) {
            if (!isDragging) return;
            e.preventDefault();

            const touch = e.originalEvent.touches[0];
            const deltaX = touch.clientX - previousMousePosition.x;
            const deltaY = touch.clientY - previousMousePosition.y;
            const now = performance.now();
            const dt = Math.max(8, now - lastDragTime);

            globeRotation.y += deltaX * 0.5;
            globeRotation.x -= deltaY * 0.5;

            globeRotation.x = Math.max(-60, Math.min(60, globeRotation.x));
            updateDragCounterRotation(globe);

            previousMousePosition = {
                x: touch.clientX,
                y: touch.clientY
            };
            const vx = (deltaX * 0.5) / dt;
            const vy = (-deltaY * 0.5) / dt;
            dragVelocity.x = dragVelocity.x * 0.8 + vx * 0.2;
            dragVelocity.y = dragVelocity.y * 0.8 + vy * 0.2;
            lastDragTime = now;
        });

        $(document).on('touchend.drag', function() {
            isDragging = false;
            $(document).off('touchmove.drag touchend.drag');
            if (!isLotteryRunning) {
                autoRotate = true;
                resumeAutoRotateFromDrag(globe);
            }
        });
    });
}

// 创建地球旋转效果（抽奖时使用）
function createGlobeEffect(remainingUsers) {
    // 先清理所有定时器和动画
    if (lotteryInterval) {
        clearInterval(lotteryInterval);
        lotteryInterval = null;
    }
    if (rotationInterval) {
        clearInterval(rotationInterval);
        rotationInterval = null;
    }
    if (autoStopTimer) {
        clearTimeout(autoStopTimer);
        autoStopTimer = null;
    }
    stopAxisShuffle();
    stopSpinEase();

    // 清空并重新创建地球
    const globeContainer = $('#globeContainer');
    globeContainer.empty();

    const globe = $('<div class="globe"></div>');
    globeContainer.append(globe);
    const axis = getRandomAxis();
    const baseDuration = 1.0;
    setGlobeSpin(globe, { axis, duration: baseDuration, delay: -Math.random() * baseDuration });

    // 在球体上放置名字
    const radius = getGlobeRadius();
    remainingUsers.forEach((user, index) => {
        const nameWrap = $('<div class="globe-name-wrap"></div>');
        const name = $('<div class="globe-name"></div>');
        name.text(user.name);

        if (user.type === 'boss') {
            name.addClass('boss-name');
        }

        // 使用斐波那契球面分布算法均匀分布名字
        const phi = Math.acos(-1 + (2 * index + 1) / remainingUsers.length);
        const theta = Math.sqrt(remainingUsers.length * Math.PI) * phi;

        const x = radius * Math.cos(theta) * Math.sin(phi);
        const y = radius * Math.sin(theta) * Math.sin(phi);
        const z = radius * Math.cos(phi);

        nameWrap.css({
            'transform': `translate3d(${x}px, ${y}px, ${z}px)`
        });

        nameWrap.append(name);
        globe.append(nameWrap);
    });

    // 重置旋转角度
    globeRotation = { x: 0, y: 0 };
    updateDragCounterRotation(globe);

    // 开始旋转并逐渐加速（但速度不会太高）
    globe.addClass('rotating');
    startAxisShuffle(globe);

    // 逐渐加速
    let speedLevel = 0;
    rotationInterval = setInterval(() => {
        speedLevel++;
        if (speedLevel <= 3) {
            const durations = [0.9, 0.7, 0.5];
            const duration = durations[speedLevel - 1] || 0.5;
            currentSpinDuration = duration;
            globe.css('--spin-duration', duration + 's');
            startAxisShuffle(globe);
        }
    }, 1000);

    // 抽奖中不显示候选人
}

// 开始抽奖
function startLottery() {
    // 添加更严格的状态检查
    if (isLotteryRunning) {
        console.log(t('lottery_running_ignore'));
        return;
    }
    if (recreateGlobeTimer) {
        clearTimeout(recreateGlobeTimer);
        recreateGlobeTimer = null;
    }

    const remainingUsers = users.filter(user => {
        const isWinner = winners.some(winner => winner.id === user.id);
        // 如果开启老板留池，老板即使中奖也不移除
        if (bossKeepInPool && user.type === 'boss') {
            return true;
        }
        return !isWinner;
    });

    if (remainingUsers.length === 0) {
        showAlert(t('no_remaining'), 'warning');
        return;
    }

    const selectedAward = getSelectedAward();
    if (!selectedAward) {
        showAlert(t('award_select_required'), 'warning');
        return;
    }
    if (selectedAward.remaining <= 0) {
        showAlert(t('award_out_of_stock'), 'warning');
        return;
    }
    if (selectedAward.perDraw <= 0) {
        showAlert(t('award_per_draw_invalid'), 'warning');
        return;
    }
    if (selectedAward.remaining < selectedAward.perDraw) {
        showAlert(t('award_not_enough_remaining', { remaining: selectedAward.remaining }), 'warning');
        return;
    }
    if (remainingUsers.length < selectedAward.perDraw) {
        showAlert(t('award_not_enough_users', { count: selectedAward.perDraw }), 'warning');
        return;
    }

    isLotteryRunning = true;
    currentDrawAwardId = selectedAward.id;
    currentLottery = null;
    $('#startBtn').prop('disabled', true);
    $('body').addClass('lottery-running');
    $('body').removeClass('winner-reveal');
    $('body').removeClass('multi-winner');
    $('#winnerPopups').empty();
    $('#lotteryNumber').text('');
    $('#lotteryName').text('');

    // 创建地球旋转效果
    createGlobeEffect(remainingUsers);

    // 根据设置决定是手动停止还是自动停止
    if (autoStopEnabled) {
        // 自动停止模式：禁用停止按钮
        $('#stopBtn').prop('disabled', true);

        // 启动自动停止定时器
        autoStopTimer = setTimeout(() => {
            stopLottery();
        }, autoStopDuration * 1000);
    } else {
        // 手动停止模式：启用停止按钮
        $('#stopBtn').prop('disabled', false);
    }
}

// 停止抽奖
function stopLottery() {
    if (!isLotteryRunning) return;

    isLotteryRunning = false;
    $('body').removeClass('lottery-running');

    // 清理所有定时器
    clearInterval(lotteryInterval);
    clearInterval(rotationInterval);
    clearTimeout(autoStopTimer);
    stopAxisShuffle();
    stopSpinEase();
    if (recreateGlobeTimer) {
        clearTimeout(recreateGlobeTimer);
        recreateGlobeTimer = null;
    }

    // 重置定时器变量
    lotteryInterval = null;
    rotationInterval = null;
    autoStopTimer = null;

    // 停止地球旋转动画
    $('.globe').removeClass('rotating accelerating');

    $('#startBtn').prop('disabled', false);
    $('#stopBtn').prop('disabled', true);

    const drawAward = awards.find((award) => award.id === currentDrawAwardId) || getSelectedAward();
    if (!drawAward) {
        showAlert(t('award_select_required'), 'warning');
        return;
    }

    const remainingUsers = users.filter(user => {
        const isWinner = winners.some(winner => winner.id === user.id);
        if (bossKeepInPool && user.type === 'boss') {
            return true;
        }
        return !isWinner;
    });

    const drawCount = Math.min(drawAward.perDraw, drawAward.remaining, remainingUsers.length);
    const pickedUsers = pickRandomUsers(remainingUsers, drawCount);
    if (pickedUsers.length === 0) {
        showAlert(t('award_no_pick'), 'warning');
        return;
    }

    const drawTime = new Date().toLocaleString(currentLang === 'en' ? 'en-US' : 'zh-CN', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        hour12: false
    });
    const drawId = Date.now() + Math.random();

    pickedUsers.forEach((winner) => {
        winners.push({
            ...winner,
            awardId: drawAward.id,
            awardName: drawAward.name,
            awardPrize: drawAward.prize,
            winTime: drawTime,
            drawId
        });
    });

    drawAward.remaining = Math.max(0, drawAward.remaining - pickedUsers.length);
    currentLottery = pickedUsers[0] || null;

    const popups = $('#winnerPopups');
    popups.empty();
    if (pickedUsers.length === 1) {
        $('body').removeClass('multi-winner');
        $('#lotteryNumber').text(pickedUsers[0].number);
        $('#lotteryName').text(pickedUsers[0].name);
        $('#lotteryName').toggleClass('boss-name', pickedUsers[0].type === 'boss');
        $('body').addClass('winner-reveal');
        const info = $('.lottery-info');
        info.removeClass('winner-pop winner-slide');
        if (info.length) {
            void info[0].offsetWidth;
        }
        info.addClass('winner-slide');
    } else {
        $('body').addClass('multi-winner').removeClass('winner-reveal');
        $('#lotteryNumber').text('');
        $('#lotteryName').text('');
        $('#lotteryName').removeClass('boss-name');
        pickedUsers.forEach((winner, index) => {
            const delay = (index * winnerMultiDelay).toFixed(2);
            popups.append(`
                <div class="winner-card winner-slide ${winner.type === 'boss' ? 'boss-name' : ''}" style="animation-delay:${delay}s;">
                    <div class="winner-name">${winner.name}</div>
                    <div class="winner-number">${winner.number}</div>
                </div>
            `);
        });
    }

    lastDraw = { drawId, awardId: drawAward.id, count: pickedUsers.length };
    $('#undoBtn').prop('disabled', false);

    // 中奖后不再弹出右上角提示

    saveData();
    updateResultTable();
    updateStats();
    renderAwards();
    updateStatsPage();

    // 重新创建球体并恢复待机旋转
    autoRotate = true;
    recreateGlobeTimer = setTimeout(() => {
        recreateGlobeTimer = null;
        if (!isLotteryRunning) {
            createGlobe();
        }
    }, 2000);
    currentDrawAwardId = null;
}

function undoLastDraw() {
    if (isLotteryRunning) {
        alert(t('running_stop_first'));
        return;
    }
    if (!lastDraw) {
        showAlert(t('undo_empty'), 'warning');
        return;
    }
    const { drawId, awardId, count } = lastDraw;
    winners = winners.filter((winner) => winner.drawId !== drawId);
    const award = awards.find((item) => item.id === awardId);
    if (award) {
        award.remaining = Math.min(award.total, award.remaining + count);
    }
    lastDraw = null;
    $('#undoBtn').prop('disabled', true);
    saveData();
    updateResultTable();
    updateStats();
    renderAwards();
    updateStatsPage();
    currentLottery = null;
    updateLotteryStatusText();
    showAlert(t('undo_success'), 'success');
}

// 重置抽奖
function resetLottery() {
    if (isLotteryRunning) {
        alert(t('running_stop_first'));
        return;
    }

    if (confirm(t('reset_confirm'))) {
        winners = [];
        currentLottery = null;
        lastDraw = null;
        $('#undoBtn').prop('disabled', true);
        awards = awards.map((award) => ({
            ...award,
            remaining: award.total
        }));
        autoRotate = true;
        saveData();
        updateResultTable();
        updateStats();
        renderAwards();
        updateStatsPage();
        createGlobe();
        $('#lotteryNumber').text('000');
        $('#lotteryName').text(t('lottery_ready'));
        showAlert(t('reset_success'), 'success');
    }
}

// 清空数据
function clearData() {
    if (isLotteryRunning) {
        alert(t('running_stop_first'));
        return;
    }

    if (confirm(t('clear_confirm'))) {
        users = [];
        winners = [];
        currentLottery = null;
        awards = [];
        selectedAwardId = null;
        lastDraw = null;
        $('#undoBtn').prop('disabled', true);
        autoRotate = true;
        saveData();
        updateTables();
        updateStats();
        renderAwards();
        updateStatsPage();
        createGlobe();
        $('#lotteryNumber').text('000');
        $('#lotteryName').text(t('lottery_ready'));
        showAlert(t('clear_success'), 'success');
    }
}

// 同步设置UI
function syncSettingsUI() {
    $('#bossKeepInPool').prop('checked', bossKeepInPool);
    $('#autoStopEnabled').prop('checked', autoStopEnabled);
    $('#autoStopDuration').val(autoStopDuration);
    $('#autoStopDurationContainer').toggle(autoStopEnabled);
    $('#winnerRevealDuration').val(winnerRevealDuration);
    $('#winnerMultiDelay').val(winnerMultiDelay);
}

// 显示提示
function showAlert(message, type) {
    const alert = $("<div class=\"alert alert-" + type + " alert-dismissible fade show\" role=\"alert\">" + message + "<button type=\"button\" class=\"btn-close\" data-bs-dismiss=\"alert\" aria-label=\"" + t('button_close') + "\"></button></div>");
    const container = $('#alertContainer');
    if (container.length) {
        container.prepend(alert);
    } else {
        $('body').prepend(alert);
    }
    setTimeout(() => {
        try {
            const el = alert[0];
            const bsAlert = new bootstrap.Alert(el);
            bsAlert.close();
        } catch (err) {
            // fallback: remove node
            alert.remove();
        }
    }, 3000);
}

// 创建粒子效果
function createParticles() {
    const lotteryArea = document.getElementById('lotteryArea');

    for (let i = 0; i < 80; i++) {
        const particle = document.createElement('div');
        particle.className = 'particle';
        particle.style.left = Math.random() * 100 + '%';
        particle.style.animationDelay = Math.random() * 4 + 's';
        particle.style.animationDuration = (Math.random() * 3 + 3) + 's';
        // 随机大小
        const size = Math.random() * 3 + 2;
        particle.style.width = size + 'px';
        particle.style.height = size + 'px';
        lotteryArea.appendChild(particle);
    }
}
