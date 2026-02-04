/**
 * 例会アンケート - アプリケーションロジック
 * 守成クラブ福岡飯塚
 */

// APIエンドポイント
// attendance GASのWebアプリURL（mode=surveyで呼び出し）
const API_ENDPOINT = 'https://script.google.com/macros/s/AKfycbxUm_qmitTYF7ckppY4LXTlVfK4MAjBi0kpSHw7NpI1LX1MWpqPzUvTXsF3x2upOStawg/exec';

/**
 * 初期化
 */
document.addEventListener('DOMContentLoaded', () => {
  initEmojiRatings();
  initFormSubmission();
});

/**
 * 絵文字評価ボタンの初期化
 */
function initEmojiRatings() {
  const ratingGroups = document.querySelectorAll('.emoji-rating');

  ratingGroups.forEach(group => {
    const fieldName = group.dataset.name;
    const hiddenInput = document.getElementById(fieldName);
    const buttons = group.querySelectorAll('.emoji-btn');

    buttons.forEach(btn => {
      btn.addEventListener('click', () => {
        // 他のボタンの選択を解除
        buttons.forEach(b => b.classList.remove('selected'));
        // このボタンを選択
        btn.classList.add('selected');
        // hidden inputに値を設定
        hiddenInput.value = btn.dataset.value;
        // エラー表示を解除
        group.classList.remove('error');
      });
    });
  });
}

/**
 * フォーム送信の初期化
 */
function initFormSubmission() {
  const form = document.getElementById('surveyForm');

  form.addEventListener('submit', async (e) => {
    e.preventDefault();

    // バリデーション
    if (!validateForm()) {
      return;
    }

    // 送信処理
    await submitSurvey();
  });
}

/**
 * フォームバリデーション
 * @returns {boolean} バリデーション結果
 */
function validateForm() {
  let isValid = true;
  let firstError = null;

  // 必須の絵文字評価をチェック
  const requiredRatings = ['meetingSatisfaction', 'clubSatisfaction'];

  requiredRatings.forEach(fieldName => {
    const hiddenInput = document.getElementById(fieldName);
    const ratingGroup = document.querySelector(`.emoji-rating[data-name="${fieldName}"]`);

    if (!hiddenInput.value) {
      ratingGroup.classList.add('error');
      if (!firstError) {
        firstError = ratingGroup;
      }
      isValid = false;
    }
  });

  // 最初のエラー項目までスクロール
  if (firstError) {
    firstError.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  return isValid;
}

/**
 * アンケート送信
 */
async function submitSurvey() {
  const submitBtn = document.getElementById('submitBtn');
  const btnText = submitBtn.querySelector('.btn-text');
  const btnLoading = submitBtn.querySelector('.btn-loading');

  // ローディング表示
  submitBtn.disabled = true;
  btnText.style.display = 'none';
  btnLoading.style.display = 'inline';

  try {
    // フォームデータを収集
    const formData = collectFormData();

    // URLパラメータから例会キーを取得
    const urlParams = new URLSearchParams(window.location.search);
    const meetingKey = urlParams.get('key') || 'unknown';
    formData.meetingKey = meetingKey;

    console.log('送信データ:', formData);

    // API送信（mode=surveyを追加）
    const payload = {
      mode: 'survey',
      ...formData
    };

    const response = await fetch(API_ENDPOINT, {
      method: 'POST',
      mode: 'cors',
      redirect: 'follow',
      headers: {
        'Content-Type': 'text/plain',
      },
      body: JSON.stringify(payload),
    });

    const result = await response.json();

    if (!result.success) {
      throw new Error(result.error || '送信に失敗しました');
    }

    // 完了画面を表示
    showThankYou();

  } catch (error) {
    console.error('送信エラー:', error);
    alert('送信に失敗しました。もう一度お試しください。');

    // ボタンを元に戻す
    submitBtn.disabled = false;
    btnText.style.display = 'inline';
    btnLoading.style.display = 'none';
  }
}

/**
 * フォームデータを収集
 * @returns {Object} フォームデータ
 */
function collectFormData() {
  // チェックボックスの値を取得するヘルパー
  const getCheckedValues = (name) => {
    const checkboxes = document.querySelectorAll(`input[name="${name}"]:checked`);
    return Array.from(checkboxes).map(cb => cb.value);
  };

  return {
    // Q1: 例会満足度
    meetingSatisfaction: parseInt(document.getElementById('meetingSatisfaction').value, 10),
    meetingGoodPoints: document.getElementById('meetingGoodPoints').value.trim(),
    meetingImprovements: document.getElementById('meetingImprovements').value.trim(),

    // Q2: システム
    attendanceSystem: getCheckedValues('attendanceSystem'),
    attendanceComment: document.getElementById('attendanceComment').value.trim(),
    salesSystem: getCheckedValues('salesSystem'),
    salesComment: document.getElementById('salesComment').value.trim(),
    otherSystemComment: document.getElementById('otherSystemComment').value.trim(),

    // Q3: 福岡飯塚満足度
    clubSatisfaction: parseInt(document.getElementById('clubSatisfaction').value, 10),
    clubGoodPoints: document.getElementById('clubGoodPoints').value.trim(),
    clubImprovements: document.getElementById('clubImprovements').value.trim(),

    // Q4: その他
    otherComments: document.getElementById('otherComments').value.trim(),

    // メタデータ
    timestamp: new Date().toISOString(),
  };
}

/**
 * 完了画面を表示
 */
function showThankYou() {
  const form = document.getElementById('surveyForm');
  const thankYou = document.getElementById('thankYou');
  const header = document.querySelector('.header');

  // フォームを非表示
  form.style.display = 'none';
  header.style.display = 'none';

  // 完了画面を表示
  thankYou.style.display = 'block';

  // ページ上部にスクロール
  window.scrollTo({ top: 0, behavior: 'smooth' });
}
