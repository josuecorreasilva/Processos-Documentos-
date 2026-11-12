(function () {
  var btn = document.querySelector(".menu-btn");
  var body = document.body;
  if (!btn) return;

  function setOpen(open) {
    body.classList.toggle("menu-aberto", open);
    btn.setAttribute("aria-expanded", open ? "true" : "false");
  }

  btn.addEventListener("click", function () {
    setOpen(!body.classList.contains("menu-aberto"));
  });

  document.addEventListener("keydown", function (e) {
    if (e.key === "Escape" && body.classList.contains("menu-aberto")) {
      setOpen(false);
      btn.focus();
    }
  });

  document.addEventListener("click", function (e) {
    if (
      body.classList.contains("menu-aberto") &&
      !e.target.closest(".lateral-nav") &&
      !e.target.closest(".menu-btn")
    ) {
      setOpen(false);
    }
  });
})();

(function () {
  var groups = document.querySelectorAll(".lnav__det");
  if (!groups.length || !window.matchMedia) return;
  var hover = window.matchMedia("(hover: hover) and (min-width: 901px)");
  groups.forEach(function (details) {
    var timer;
    details.addEventListener("mouseenter", function () {
      if (!hover.matches) return;
      clearTimeout(timer);
      details.open = true;
    });
    details.addEventListener("mouseleave", function () {
      if (!hover.matches || details.hasAttribute("data-fixo")) return;
      timer = setTimeout(function () {
        details.open = false;
      }, 280);
    });
  });
})();
