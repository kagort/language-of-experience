// Sidebar toggle for mobile
(function () {
  var toggle  = document.getElementById('sidebar-toggle');
  var sidebar = document.getElementById('sidebar');
  var overlay = document.getElementById('sidebar-overlay');

  function open()  { sidebar.classList.add('is-open');  overlay.classList.add('is-open');  }
  function close() { sidebar.classList.remove('is-open'); overlay.classList.remove('is-open'); }

  if (toggle)  toggle.addEventListener('click', function () { sidebar.classList.contains('is-open') ? close() : open(); });
  if (overlay) overlay.addEventListener('click', close);

  // Close on Escape
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') close(); });

  // Mark active nav link as "visited" after navigation
  var activeLink = document.querySelector('.section-item.is-active .section-link');
  if (activeLink) activeLink.scrollIntoView({ block: 'nearest' });
})();
