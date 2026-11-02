// script.js - Pham Bao Phong - 1822040447
document.addEventListener('DOMContentLoaded', () => {
  const btnCheckStatus = document.getElementById('btnCheckStatus');
  const btnToggleTheme = document.getElementById('btnToggleTheme');
  const statusResult = document.getElementById('statusResult');
  const liveClock = document.getElementById('liveClock');

  // Update live clock
  function updateClock() {
    const now = new Date();
    liveClock.textContent = `Thời gian hệ thống: ${now.toLocaleDateString('vi-VN')} ${now.toLocaleTimeString('vi-VN')}`;
  }
  updateClock();
  setInterval(updateClock, 1000);

  // Status check action
  btnCheckStatus.addEventListener('click', () => {
    statusResult.classList.remove('hidden');
    statusResult.innerHTML = `
      <strong>✓ Hệ thống hoạt động hoàn hảo!</strong><br>
      • Sinh viên: Phạm Bảo Phong (1822040447)<br>
      • Pipeline CI/CD: Sẵn sàng tự động kiểm thử và deploy<br>
      • Thời điểm xác thực: ${new Date().toLocaleTimeString('vi-VN')}
    `;
  });

  // Toggle theme action
  btnToggleTheme.addEventListener('click', () => {
    const currentTheme = document.documentElement.getAttribute('data-theme');
    const newTheme = currentTheme === 'light' ? 'dark' : 'light';
    document.documentElement.setAttribute('data-theme', newTheme);
  });
});
