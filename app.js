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

// 鼠标拖动旋转相关变量
let isDragging = false;
let previousMousePosition = { x: 0, y: 0 };
let globeRotation = { x: 0, y: 0 };
let autoRotate = true; // 是否自动旋转
let lastDragTime = 0;
let dragVelocity = { x: 0, y: 0 };

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
    $('#addForm').submit(addUser);
    $('#userExcel').change(importUsers);
    $('#bossExcel').change(importBosses);
    $('#saveEditBtn').click(saveEdit);
    $('#clearDataBtn').click(clearData);
    $('#bossKeepInPool').change(function() {
        bossKeepInPool = $(this).is(':checked');
        localStorage.setItem('bossKeepInPool', bossKeepInPool);
        showAlert('设置已保存！', 'success');
    });
    $('#autoStopEnabled').change(function() {
        autoStopEnabled = $(this).is(':checked');
        localStorage.setItem('autoStopEnabled', autoStopEnabled);
        $('#autoStopDurationContainer').toggle(autoStopEnabled);
        showAlert('设置已保存！', 'success');
    });
    $('#autoStopDuration').change(function() {
        autoStopDuration = parseInt($(this).val()) || 5;
        localStorage.setItem('autoStopDuration', autoStopDuration);
    });

    $('#drawerToggle').on('click', function() {
        $('body').toggleClass('drawer-open');
        const isOpen = $('body').hasClass('drawer-open');
        $(this).attr('aria-label', isOpen ? '收起抽奖结果' : '展开抽奖结果');
    });
    $('#drawerClose').on('click', function() {
        $('body').removeClass('drawer-open');
    });
    $('#drawerBackdrop').on('click', function() {
        $('body').removeClass('drawer-open');
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
});

// 加载数据
function loadData() {
    users = JSON.parse(localStorage.getItem('lotteryUsers')) || [];
    winners = JSON.parse(localStorage.getItem('lotteryWinners')) || [];
    bossKeepInPool = localStorage.getItem('bossKeepInPool') !== 'false'; // 默认为true
    autoStopEnabled = localStorage.getItem('autoStopEnabled') === 'true';
    autoStopDuration = parseInt(localStorage.getItem('autoStopDuration')) || 5;
}

// 保存数据
function saveData() {
    localStorage.setItem('lotteryUsers', JSON.stringify(users));
    localStorage.setItem('lotteryWinners', JSON.stringify(winners));
}

// 更新表格
function updateTables() {
    updateUserTable();
    updateResultTable();
}

// 更新人员表格
function updateUserTable() {
    const tbody = $('#userTable');
    tbody.empty();

    users.forEach((user, index) => {
        const tr = $('<tr></tr>');
        tr.append($('<td></td>').text(user.name));
        tr.append($('<td></td>').text(user.number));

        const typeBadge = user.type === 'boss' 
            ? $('<span class="badge badge-boss"><i class="fa fa-star"></i> 老板</span>')
            : $('<span class="badge bg-secondary">员工</span>');
        tr.append($('<td></td>').append(typeBadge));

        const actions = $('<td></td>');
        actions.append($('<button class="btn btn-sm btn-primary me-1" onclick="editUser(' + index + ')"><i class="fa fa-pencil"></i></button>'));
        actions.append($('<button class="btn btn-sm btn-danger" onclick="deleteUser(' + index + ')"><i class="fa fa-trash"></i></button>'));
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
            tr.append($('<td></td>').text(winner.name));
            tr.append($('<td></td>').text(winner.number));

            const typeBadge = winner.type === 'boss'
                ? $('<span class="badge badge-boss"><i class="fa fa-star"></i> 老板</span>')
                : $('<span class="badge bg-secondary">员工</span>');
            tr.append($('<td></td>').append(typeBadge));

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
        const text = isFullscreen ? '退出全屏' : '全屏展示';
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

// 添加用户
function addUser(e) {
    e.preventDefault();

    const name = $('#addName').val();
    const number = $('#addNumber').val();
    const type = $('#addType').val();

    // 检查号码是否重复
    if (users.some(user => user.number === number)) {
        alert('该号码已存在！');
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
    showAlert('添加成功！', 'success');
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
        alert('该号码已存在！');
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
    showAlert('更新成功！', 'success');
}

// 删除用户
function deleteUser(index) {
    if (confirm('确定要删除该用户吗？')) {
        users.splice(index, 1);
        saveData();
        updateTables();
        updateStats();
        updateNameWall();
        showAlert('删除成功！', 'success');
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
        showAlert('Excel文件为空！', 'warning');
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
        showAlert('无法识别Excel列名，请确保包含"姓名"和"号码"列！', 'danger');
        return;
    }

    if (nameKey === numberKey) {
        showAlert('姓名列和号码列不能是同一列！请检查Excel文件格式。', 'danger');
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
        showAlert('Excel中没有有效的数据！请检查姓名和号码列是否都有值。', 'warning');
        return;
    }

    saveData();
    updateTables();
    updateStats();
    updateNameWall();

    let message = `导入成功！新增 ${importedCount} 人`;
    if (duplicateCount > 0) {
        message += `，跳过重复 ${duplicateCount} 人`;
    }
    if (skippedEmpty > 0) {
        message += `，跳过空数据 ${skippedEmpty} 行`;
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

    // 随机选择中奖者
    lotteryInterval = setInterval(() => {
        const randomIndex = Math.floor(Math.random() * remainingUsers.length);
        currentLottery = remainingUsers[randomIndex];
    }, 100);
}

// 开始抽奖
function startLottery() {
    // 添加更严格的状态检查
    if (isLotteryRunning) {
        console.log('抽奖正在进行中，忽略重复点击');
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
        alert('没有剩余抽奖人员！');
        return;
    }

    isLotteryRunning = true;
    $('#startBtn').prop('disabled', true);
    $('body').addClass('lottery-running');
    $('body').removeClass('winner-reveal');
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

    if (currentLottery) {
        $('#lotteryNumber').text(currentLottery.number);
        $('#lotteryName').text(currentLottery.name);
        $('body').addClass('winner-reveal');
        const info = $('.lottery-info');
        info.removeClass('winner-pop');
        if (info.length) {
            void info[0].offsetWidth;
        }
        info.addClass('winner-pop');
        // 添加中奖时间
        const winnerWithTime = {
            ...currentLottery,
            winTime: new Date().toLocaleString('zh-CN', {
                year: 'numeric',
                month: '2-digit',
                day: '2-digit',
                hour: '2-digit',
                minute: '2-digit',
                second: '2-digit',
                hour12: false
            })
        };
        winners.push(winnerWithTime);

        // 根据设置显示不同的提示
        if (currentLottery.type === 'boss') {
            if (bossKeepInPool) {
                showAlert(`恭喜 ${currentLottery.name} (${currentLottery.number}) 中奖！老板可以继续参与抽奖。`, 'info');
            } else {
                showAlert(`恭喜 ${currentLottery.name} (${currentLottery.number}) 中奖！`, 'success');
            }
        } else {
            showAlert(`恭喜 ${currentLottery.name} (${currentLottery.number}) 中奖！`, 'success');
        }

        saveData();
        updateResultTable();
        updateStats();

        // 重新创建球体并恢复待机旋转
        autoRotate = true;
        recreateGlobeTimer = setTimeout(() => {
            recreateGlobeTimer = null;
            if (!isLotteryRunning) {
                createGlobe();
            }
        }, 2000);
    }
}

// 重置抽奖
function resetLottery() {
    if (isLotteryRunning) {
        alert('抽奖进行中，请先停止抽奖！');
        return;
    }

    if (confirm('确定要重置所有抽奖结果吗？')) {
        winners = [];
        currentLottery = null;
        autoRotate = true;
        saveData();
        updateResultTable();
        updateStats();
        createGlobe();
        $('#lotteryNumber').text('000');
        $('#lotteryName').text('准备抽奖');
        showAlert('抽奖结果已重置！', 'success');
    }
}

// 清空数据
function clearData() {
    if (isLotteryRunning) {
        alert('抽奖进行中，请先停止抽奖！');
        return;
    }

    if (confirm('确定要清空所有数据吗？此操作不可恢复！')) {
        users = [];
        winners = [];
        currentLottery = null;
        autoRotate = true;
        saveData();
        updateTables();
        updateStats();
        createGlobe();
        $('#lotteryNumber').text('000');
        $('#lotteryName').text('准备抽奖');
        showAlert('所有数据已清空！', 'success');
    }
}

// 同步设置UI
function syncSettingsUI() {
    $('#bossKeepInPool').prop('checked', bossKeepInPool);
    $('#autoStopEnabled').prop('checked', autoStopEnabled);
    $('#autoStopDuration').val(autoStopDuration);
    $('#autoStopDurationContainer').toggle(autoStopEnabled);
}

// 显示提示
function showAlert(message, type) {
    const alert = $("<div class=\"alert alert-" + type + " alert-dismissible fade show\" role=\"alert\">" + message + "<button type=\"button\" class=\"btn-close\" data-bs-dismiss=\"alert\" aria-label=\"Close\"></button></div>");
    $('.container').prepend(alert);
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
