const searchInput = document.querySelector("[data-macro-search]");
const categorySelect = document.querySelector("[data-category-filter]");
const cards = Array.from(document.querySelectorAll(".macro-card"));
const count = document.querySelector("[data-result-count]");

function applyFilters() {
  const query = (searchInput?.value || "").trim().toLowerCase();
  const category = categorySelect?.value || "All";
  let visible = 0;
  cards.forEach((card) => {
    const matchesQuery = !query || card.dataset.search.includes(query);
    const matchesCategory = category === "All" || card.dataset.category === category;
    const show = matchesQuery && matchesCategory;
    card.hidden = !show;
    if (show) visible += 1;
  });
  if (count) count.textContent = String(visible);
}

searchInput?.addEventListener("input", applyFilters);
categorySelect?.addEventListener("change", applyFilters);
applyFilters();
