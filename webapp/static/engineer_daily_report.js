(function () {
  const dateEl = document.getElementById("reportDate");
  if (dateEl && !dateEl.value) {
    dateEl.value = new Date().toISOString().slice(0, 10);
  }
})();
