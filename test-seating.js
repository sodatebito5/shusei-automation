const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({ headless: true });

  // iPhone 12 Pro サイズ
  const context = await browser.newContext({
    viewport: { width: 390, height: 844 },
    deviceScaleFactor: 3,
    isMobile: true,
    hasTouch: true,
  });

  const page = await context.newPage();

  console.log('ダッシュボードを開いています...');
  await page.goto('https://script.google.com/macros/s/AKfycbz5j-qZV2RW5nU2PoYEQYUQKooGSWtboHMPOgjSIQI/dev', {
    waitUntil: 'networkidle',
    timeout: 60000
  });

  console.log('ページ読み込み完了、8秒待機...');
  await page.waitForTimeout(8000);

  // 初期スクリーンショット
  await page.screenshot({ path: 'test-results/seating-1-initial.png', fullPage: false });
  console.log('初期画面スクリーンショット完了');

  // 例会当日タブをクリック（テキストで検索）
  console.log('例会当日タブをクリック...');
  try {
    await page.click('text=例会当日', { timeout: 5000 });
    await page.waitForTimeout(2000);
    await page.screenshot({ path: 'test-results/seating-2-eventday.png', fullPage: false });
    console.log('例会当日タブスクリーンショット完了');
  } catch (e) {
    console.log('例会当日タブが見つかりません:', e.message);
  }

  // 配席表ボタンをクリック
  console.log('配席表ボタンをクリック...');
  try {
    await page.click('text=配席表', { timeout: 5000 });
    await page.waitForTimeout(4000);
    await page.screenshot({ path: 'test-results/seating-3-modal.png', fullPage: false });
    console.log('配席表モーダルスクリーンショット完了');

    // モーダル内のコンテンツサイズを確認
    const modalInfo = await page.evaluate(() => {
      const modal = document.querySelector('.seating-modal-content');
      const container = document.querySelector('.seating-chart-container');
      const body = document.querySelector('.seating-modal-body');

      if (!modal || !container) return { error: 'モーダルが見つかりません' };

      return {
        viewport: { width: window.innerWidth, height: window.innerHeight },
        modal: { width: modal.offsetWidth, height: modal.offsetHeight },
        container: { width: container.scrollWidth, height: container.scrollHeight },
        body: { width: body.clientWidth, height: body.clientHeight },
        overflow: container.scrollWidth > body.clientWidth ? 'はみ出し' : 'OK'
      };
    });

    console.log('サイズ情報:', JSON.stringify(modalInfo, null, 2));
  } catch (e) {
    console.log('配席表ボタンが見つかりません:', e.message);
  }

  await browser.close();
  console.log('完了');
})();
