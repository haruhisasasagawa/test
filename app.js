let reminders = [];

// ページ読み込み時の初期化
document.addEventListener('DOMContentLoaded', () => {
    loadReminders();
    displayReminders();
    requestNotificationPermission();

    // 1秒ごとにリマインダーをチェック
    setInterval(checkReminders, 1000);

    // デフォルトの日時を設定（現在時刻の1時間後）
    const now = new Date();
    now.setHours(now.getHours() + 1);
    now.setMinutes(0);
    const defaultTime = now.toISOString().slice(0, 16);
    document.getElementById('reminderTime').value = defaultTime;
});

// 通知の許可をリクエスト
function requestNotificationPermission() {
    if ('Notification' in window && Notification.permission === 'default') {
        Notification.requestPermission();
    }
}

// ローカルストレージからリマインダーを読み込む
function loadReminders() {
    const stored = localStorage.getItem('reminders');
    if (stored) {
        reminders = JSON.parse(stored);
    }
}

// ローカルストレージにリマインダーを保存
function saveReminders() {
    localStorage.setItem('reminders', JSON.stringify(reminders));
}

// リマインダーを追加
function addReminder() {
    const title = document.getElementById('reminderTitle').value.trim();
    const time = document.getElementById('reminderTime').value;

    if (!title) {
        alert('タイトルを入力してください');
        return;
    }

    if (!time) {
        alert('日時を選択してください');
        return;
    }

    const reminderTime = new Date(time);
    const now = new Date();

    if (reminderTime <= now) {
        alert('未来の日時を選択してください');
        return;
    }

    const reminder = {
        id: Date.now(),
        title: title,
        time: time,
        notified: false
    };

    reminders.push(reminder);
    saveReminders();
    displayReminders();

    // 入力フィールドをクリア
    document.getElementById('reminderTitle').value = '';

    // 次のデフォルト時刻を設定
    const nextTime = new Date();
    nextTime.setHours(nextTime.getHours() + 1);
    nextTime.setMinutes(0);
    document.getElementById('reminderTime').value = nextTime.toISOString().slice(0, 16);
}

// リマインダーを削除
function deleteReminder(id) {
    reminders = reminders.filter(reminder => reminder.id !== id);
    saveReminders();
    displayReminders();
}

// リマインダーを表示
function displayReminders() {
    const listElement = document.getElementById('remindersList');
    const emptyMessage = document.getElementById('emptyMessage');

    if (reminders.length === 0) {
        listElement.innerHTML = '';
        emptyMessage.classList.remove('hidden');
        return;
    }

    emptyMessage.classList.add('hidden');

    // 時刻順にソート（近い順）
    const sortedReminders = [...reminders].sort((a, b) => {
        return new Date(a.time) - new Date(b.time);
    });

    listElement.innerHTML = sortedReminders.map(reminder => {
        const reminderTime = new Date(reminder.time);
        const now = new Date();
        const isOverdue = reminderTime < now;

        const formattedTime = formatDateTime(reminderTime);
        const timeStatus = isOverdue ? '期限切れ' : formattedTime;

        return `
            <div class="reminder-item ${isOverdue ? 'overdue' : ''}">
                <div class="reminder-info">
                    <div class="reminder-title">${escapeHtml(reminder.title)}</div>
                    <div class="reminder-time ${isOverdue ? 'overdue-text' : ''}">${timeStatus}</div>
                </div>
                <button class="delete-button" onclick="deleteReminder(${reminder.id})">削除</button>
            </div>
        `;
    }).join('');
}

// 日時をフォーマット
function formatDateTime(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');

    return `${year}年${month}月${day}日 ${hours}:${minutes}`;
}

// HTMLエスケープ
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// リマインダーをチェックして通知
function checkReminders() {
    const now = new Date();

    reminders.forEach(reminder => {
        if (!reminder.notified) {
            const reminderTime = new Date(reminder.time);

            // リマインダーの時刻になったら通知
            if (now >= reminderTime) {
                showNotification(reminder);
                reminder.notified = true;
                saveReminders();
                displayReminders();
            }
        }
    });
}

// 通知を表示
function showNotification(reminder) {
    // ブラウザ通知
    if ('Notification' in window && Notification.permission === 'granted') {
        new Notification('🔔 リマインダー', {
            body: reminder.title,
            icon: '📝',
            requireInteraction: true
        });
    }

    // アラート（フォールバック）
    alert(`🔔 リマインダー\n\n${reminder.title}`);
}

// Enterキーで追加
document.addEventListener('DOMContentLoaded', () => {
    document.getElementById('reminderTitle').addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            addReminder();
        }
    });
});
