let blockCounter = 0;
let pleasureCounter = 0;

Office.onReady(function () {
  const insertBtn = document.getElementById('insertBtn');
  const previewBtn = document.getElementById('previewBtn');
  const addImageBlockBtn = document.getElementById('addImageBlock');
  const addPleasureBtn = document.getElementById('addPleasure');

  if (insertBtn) insertBtn.addEventListener('click', buildInOutlook);
  if (previewBtn) previewBtn.addEventListener('click', preview);
  if (addImageBlockBtn) {
    addImageBlockBtn.addEventListener('click', function () {
      addImageBlock();
    });
  }
  if (addPleasureBtn) {
    addPleasureBtn.addEventListener('click', function () {
      addPleasureRow();
    });
  }

  setupHero('hero1');
  setupHero('hero2');

  document.querySelectorAll('[data-url]').forEach(function (button) {
    button.addEventListener('click', function () {
      const key = button.getAttribute('data-url');
      const input = document.getElementById(key + 'Url');
      if (input) input.focus();
      setStatus('Paste the image URL into the field above.');
    });
  });

  document.querySelectorAll('[data-drive]').forEach(function (button) {
    button.addEventListener('click', function () {
      openDriveChooser(button.getAttribute('data-drive'));
    });
  });

  addPleasureRow('Coffee', '');
  addPleasureRow('Art', '');
  addPleasureRow('Object', '');
  addImageBlock();
});

function value(id) {
  const element = document.getElementById(id);
  return element ? element.value.trim() : '';
}

function escapeHtml
