const { test, expect } = require('@playwright/test');
const { openApp, uploadFile, waitForWindow, executeSQL } = require('../helpers');

test('vertical split: move bottom to right fills full height', async ({ page }) => {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');
  await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
  await waitForWindow(page, 't2');

  // Split vertically (top/bottom)
  await page.evaluate(() => {
    const w = app._test.windows;
    app._test.mergeWindowsAsSplit(w[0].id, w[1].id, 'bottom');
  });

  // Move bottom window to the right of the top one
  const result = await page.evaluate(() => {
    const w = app._test.windows;
    const dock = app._test.dockContainers[0];
    const w1 = w[1];
    const sourceLeaf = w1.dockLeaf;
    function findLeaves(node) {
      if (node.type === 'leaf') return [node];
      return node.children.flatMap(c => findLeaves(c));
    }
    const targetLeaf = findLeaves(dock.root).find(l => l !== sourceLeaf);
    app._test.removeTabFromLeaf(w1.id, sourceLeaf, dock, true);
    w1.el.classList.add('docked');
    app._test.dockWindowAsSplit(w1.id, targetLeaf, dock, 'right');
    app._test.cleanupAfterTabRemoval(sourceLeaf, dock);

    const dockEl = dock.el;
    const dockHeight = dockEl.offsetHeight;
    const rootEl = dockEl.querySelector('.dock-root');
    const rootHeight = rootEl.offsetHeight;
    const splitEl = rootEl.firstElementChild;
    const splitHeight = splitEl ? splitEl.offsetHeight : 0;
    const visibleBodies = document.querySelectorAll('.dock-leaf-content .win-body:not([style*="display: none"])').length;
    const rootFlex = dock.root.flex;

    return { dockHeight, rootHeight, splitHeight, visibleBodies, rootFlex,
             rootFillsContainer: rootHeight >= dockHeight - 2,
             splitFillsRoot: splitHeight >= rootHeight - 2 };
  });
  console.log('After move bottom to right:', JSON.stringify(result));
  expect(result.visibleBodies).toBe(2);
  expect(result.rootFlex).toBeUndefined();
  expect(result.splitFillsRoot).toBe(true);
});

test('same-dock move: right, down, right', async ({ page }) => {
  await openApp(page);
  await uploadFile(page, 'sample1.csv');
  await waitForWindow(page, 'sample1');
  await executeSQL(page, "SELECT * INTO [t2] FROM [sample1]");
  await waitForWindow(page, 't2');
  await executeSQL(page, "SELECT * INTO [t3] FROM [sample1]");
  await waitForWindow(page, 't3');

  // Tab all 3 windows together
  await page.evaluate(() => {
    const w = app._test.windows;
    app._test.mergeWindowsIntoTabs(w[0].id, w[1].id);
    const dock = app._test.dockContainers[0];
    const leaf = dock.root;
    app._test.dockWindowAsTab(w[2].id, leaf, dock);
    w[2].el.classList.add('docked');
  });

  // Move 1: move w[2] to the right (split)
  const r2 = await page.evaluate(() => {
    const w = app._test.windows;
    const dock = app._test.dockContainers[0];
    const w2 = w[2];
    const leaf = w2.dockLeaf;
    app._test.removeTabFromLeaf(w2.id, leaf, dock, true);
    w2.el.classList.add('docked');
    app._test.dockWindowAsSplit(w2.id, leaf, dock, 'right');
    app._test.cleanupAfterTabRemoval(leaf, dock);
    const visibleBodies = document.querySelectorAll('.dock-leaf-content .win-body:not([style*="display: none"])');
    return { visibleBodies: visibleBodies.length, docks: app._test.dockContainers.length };
  });
  console.log('After move right:', JSON.stringify(r2));
  expect(r2.visibleBodies).toBeGreaterThanOrEqual(2);

  // Move 2: move w[2] down (onto the other leaf)
  const r3 = await page.evaluate(() => {
    const w = app._test.windows;
    const dock = app._test.dockContainers[0];
    const w2 = w[2];
    const sourceLeaf = w2.dockLeaf;
    function findLeaves(node) {
      if (node.type === 'leaf') return [node];
      return node.children.flatMap(c => findLeaves(c));
    }
    const leaves = findLeaves(dock.root);
    const targetLeaf = leaves.find(l => l !== sourceLeaf && l.tabs.length > 0);

    const info = {
      sourceLeafTabs: sourceLeaf.tabs.slice(),
      targetLeafTabs: targetLeaf.tabs.slice(),
      targetActiveTab: targetLeaf.activeTab,
    };

    app._test.removeTabFromLeaf(w2.id, sourceLeaf, dock, true);

    const afterRemove = {
      w0body: !!w[0].el.querySelector('.win-body'),
      w1body: !!w[1].el.querySelector('.win-body'),
      w2body: !!w[2].el.querySelector('.win-body'),
    };

    w2.el.classList.add('docked');
    app._test.dockWindowAsSplit(w2.id, targetLeaf, dock, 'bottom');

    const afterSplit = {
      visibleBodies: document.querySelectorAll('.dock-leaf-content .win-body:not([style*="display: none"])').length,
      allBodiesInDock: document.querySelectorAll('.dock-leaf-content .win-body').length,
      w0body: !!w[0].el.querySelector('.win-body'),
      w1body: !!w[1].el.querySelector('.win-body'),
      w2body: !!w[2].el.querySelector('.win-body'),
    };

    app._test.cleanupAfterTabRemoval(sourceLeaf, dock);

    const afterCleanup = {
      visibleBodies: document.querySelectorAll('.dock-leaf-content .win-body:not([style*="display: none"])').length,
      allBodiesInDock: document.querySelectorAll('.dock-leaf-content .win-body').length,
      w0body: !!w[0].el.querySelector('.win-body'),
      w1body: !!w[1].el.querySelector('.win-body'),
      w2body: !!w[2].el.querySelector('.win-body'),
      docks: app._test.dockContainers.length,
      rootType: dock.root ? dock.root.type : 'null',
    };

    return { info, afterRemove, afterSplit, afterCleanup };
  });
  console.log('After move down:', JSON.stringify(r3, null, 2));
  expect(r3.afterCleanup.visibleBodies).toBeGreaterThanOrEqual(2);

  // Move 3: move w[2] right again
  const r4 = await page.evaluate(() => {
    const w = app._test.windows;
    if (!app._test.dockContainers.length) return { error: 'no dock containers' };
    const dock = app._test.dockContainers[0];
    const w2 = w[2];
    if (!w2.dockLeaf) return { error: 'w2 not docked', dockId: w2.dockId };
    const sourceLeaf = w2.dockLeaf;
    function findLeaves(node) {
      if (node.type === 'leaf') return [node];
      return node.children.flatMap(c => findLeaves(c));
    }
    const leaves = findLeaves(dock.root);
    const targetLeaf = leaves.find(l => l !== sourceLeaf && l.tabs.length > 0);
    if (!targetLeaf) return { error: 'no target leaf', leaves: leaves.length, leafTabs: leaves.map(l => l.tabs.length) };
    app._test.removeTabFromLeaf(w2.id, sourceLeaf, dock, true);
    w2.el.classList.add('docked');
    app._test.dockWindowAsSplit(w2.id, targetLeaf, dock, 'right');
    app._test.cleanupAfterTabRemoval(sourceLeaf, dock);
    const visibleBodies = document.querySelectorAll('.dock-leaf-content .win-body:not([style*="display: none"])');
    const allBodies = document.querySelectorAll('.dock-leaf-content .win-body');
    return {
      visibleBodies: visibleBodies.length,
      allBodies: allBodies.length,
      docks: app._test.dockContainers.length,
    };
  });
  console.log('After move right again:', JSON.stringify(r4));
  expect(r4.error).toBeUndefined();
  expect(r4.docks).toBe(1);
  expect(r4.visibleBodies).toBeGreaterThanOrEqual(2);
});
