const { chromium } = require('playwright');

const DASHBOARD_URL = 'https://script.google.com/macros/s/AKfycbzgVIzNjfW_UHihZS5bwrWM7xix0U4dnodZSlq7nPC8eGXGu_Fj6haCzivxiARVDPGL/exec';

(async () => {
  console.log('ブラウザを起動しています...');

  const browser = await chromium.launch({
    headless: false,
    slowMo: 300
  });

  const context = await browser.newContext();
  const page = await context.newPage();

  // コンソールログを出力
  page.on('console', msg => {
    const text = msg.text();
    if (!text.includes('color:') && !text.includes('font-size') && !text.includes('Unrecognized feature')) {
      console.log(`[CONSOLE ${msg.type()}]: ${text}`);
    }
  });

  // エラーをキャッチ
  page.on('pageerror', err => {
    console.log(`[PAGE ERROR]: ${err.message}`);
  });

  try {
    console.log('\n=== ダッシュボードを開いています ===');
    await page.goto(DASHBOARD_URL, { waitUntil: 'load', timeout: 120000 });

    // Googleログインページかどうかを確認
    const url = page.url();
    if (url.includes('accounts.google.com')) {
      console.log('\n*** Googleログインが必要です ***');
      await page.waitForURL('**/exec**', { timeout: 120000 });
    }

    // ページの読み込み待機
    await page.waitForTimeout(5000);
    console.log('ページ読み込み完了\n');

    // iframe内のフレームを取得
    console.log('=== iframe内のコンテンツを探しています ===');
    const frames = page.frames();
    console.log(`フレーム数: ${frames.length}`);

    let targetFrame = null;
    for (const frame of frames) {
      const frameUrl = frame.url();
      console.log(`フレームURL: ${frameUrl.substring(0, 80)}...`);
      if (frameUrl.includes('userCodeAppPanel')) {
        targetFrame = frame;
        console.log('→ ターゲットフレームを発見！');
        break;
      }
    }

    if (!targetFrame) {
      // メインページで試す
      console.log('iframeが見つからないため、メインページで操作します');
      targetFrame = page;
    }

    // 会員管理タブを探す
    console.log('\n=== 会員管理タブを探しています ===');

    // タブボタンを探す
    const memberTab = await targetFrame.locator('button:has-text("会員管理"), a:has-text("会員管理"), [data-tab="member"], .tab:has-text("会員管理")').first();

    if (await memberTab.count() > 0) {
      await memberTab.click();
      console.log('会員管理タブをクリックしました');

      await targetFrame.waitForTimeout(2000);

      // パスワード入力待ち
      const passwordInput = await targetFrame.locator('#memberEditPasswordInput');
      if (await passwordInput.count() > 0) {
        console.log('パスワード入力欄が表示されました');
      }
    } else {
      console.log('会員管理タブが見つかりません');

      // ページ内のすべてのボタンとリンクを表示
      const buttons = await targetFrame.locator('button, a, .tab').all();
      console.log(`\n見つかった要素数: ${buttons.length}`);
      for (let i = 0; i < Math.min(buttons.length, 20); i++) {
        const text = await buttons[i].textContent();
        if (text && text.trim()) {
          console.log(`  - "${text.trim().substring(0, 30)}"`);
        }
      }
    }

    console.log('\n*** 手動操作に切り替えます ***');
    console.log('1. パスワードを入力して認証');
    console.log('2. 会員を選んで「編集」をクリック');
    console.log('3. 何か変更して「保存」をクリック');
    console.log('\nコンソールログを監視しています（5分間）...\n');

    // 5分間待機（手動操作用）
    await page.waitForTimeout(300000);

  } catch (e) {
    console.error('\nエラー発生:', e.message);
    await page.screenshot({ path: 'debug-error.png' });
    console.log('スクリーンショットを保存しました: debug-error.png');
  }

  console.log('\nブラウザを閉じています...');
  await browser.close();
})();
