(function () {
  "use strict";

  var tabs = Array.prototype.slice.call(document.querySelectorAll(".tab"));
  var panels = Array.prototype.slice.call(document.querySelectorAll(".panel"));
  var targetButtons = Array.prototype.slice.call(document.querySelectorAll("[data-target]"));
  var workspace = document.getElementById("workspace");
  var toggleEmbedBtn = document.querySelector('[data-action="toggle-stitch-embed"]');
  var stitchEmbedWrap = document.getElementById("stitch-embed-wrap");
  var stitchEmbed = document.getElementById("stitch-embed");
  var stitchResizeTimer = null;
  var yearLabel = document.getElementById("year");

  function setActivePanel(target) {
    var hasTargetPanel = panels.some(function (panel) {
      return panel.id === "panel-" + target;
    });
    if (!hasTargetPanel) {
      return;
    }

    tabs.forEach(function (tab) {
      tab.classList.toggle("active", tab.getAttribute("data-target") === target);
    });

    panels.forEach(function (panel) {
      panel.classList.toggle("active", panel.id === "panel-" + target);
    });
  }

  function scrollToWorkspace() {
    if (!workspace) {
      return;
    }
    workspace.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  function syncStitchEmbedHeight() {
    if (!stitchEmbed) {
      return;
    }
    try {
      var doc = stitchEmbed.contentDocument;
      if (!doc || !doc.body) {
        return;
      }
      var height = Math.max(
        doc.body.scrollHeight || 0,
        doc.documentElement ? doc.documentElement.scrollHeight : 0
      );
      if (height > 0) {
        stitchEmbed.style.height = String(height + 20) + "px";
      }
    } catch (error) {
      // If access is restricted in a given browser context, keep the default height.
    }
  }

  function fallbackCopyText(text) {
    var textarea = document.createElement("textarea");
    textarea.value = text;
    textarea.setAttribute("readonly", "readonly");
    textarea.style.position = "fixed";
    textarea.style.top = "-1000px";
    textarea.style.opacity = "0";
    document.body.appendChild(textarea);
    textarea.focus();
    textarea.select();

    var ok = false;
    try {
      ok = document.execCommand("copy");
    } catch (error) {
      ok = false;
    }

    document.body.removeChild(textarea);
    return ok;
  }

  function copyText(text) {
    if (navigator.clipboard && window.isSecureContext) {
      return navigator.clipboard.writeText(text);
    }

    return new Promise(function (resolve, reject) {
      if (fallbackCopyText(text)) {
        resolve();
      } else {
        reject(new Error("Copy failed"));
      }
    });
  }

  tabs.forEach(function (tab) {
    tab.addEventListener("click", function () {
      var target = tab.getAttribute("data-target");
      setActivePanel(target);
    });
  });

  targetButtons.forEach(function (button) {
    button.addEventListener("click", function () {
      var target = button.getAttribute("data-target");
      if (!target) {
        return;
      }
      setActivePanel(target);
      scrollToWorkspace();
    });
  });

  document.querySelectorAll('[data-action="open-workspace"]').forEach(function (button) {
    button.addEventListener("click", scrollToWorkspace);
  });

  if (toggleEmbedBtn && stitchEmbedWrap) {
    toggleEmbedBtn.addEventListener("click", function () {
      var isHidden = stitchEmbedWrap.hasAttribute("hidden");
      if (isHidden) {
        stitchEmbedWrap.removeAttribute("hidden");
        toggleEmbedBtn.textContent = "Hide embedded tool";
        syncStitchEmbedHeight();
      } else {
        stitchEmbedWrap.setAttribute("hidden", "hidden");
        toggleEmbedBtn.textContent = "Show embedded tool";
      }
    });
  }

  if (stitchEmbed) {
    stitchEmbed.addEventListener("load", function () {
      syncStitchEmbedHeight();
      if (stitchResizeTimer) {
        window.clearInterval(stitchResizeTimer);
      }
      stitchResizeTimer = window.setInterval(syncStitchEmbedHeight, 1200);
    });
    window.addEventListener("resize", syncStitchEmbedHeight);
  }

  document.querySelectorAll("[data-copy-target]").forEach(function (button) {
    button.addEventListener("click", function () {
      var targetId = button.getAttribute("data-copy-target");
      var source = document.getElementById(targetId);
      if (!source) {
        return;
      }
      var text = source.innerText.trim();
      if (!text) {
        return;
      }

      copyText(text).then(function () {
        var original = button.textContent;
        button.textContent = "Copied";
        button.disabled = true;
        window.setTimeout(function () {
          button.textContent = original;
          button.disabled = false;
        }, 1400);
      }).catch(function () {
        button.textContent = "Copy failed";
      });
    });
  });

  if (yearLabel) {
    yearLabel.textContent = "Dashboard ready - " + new Date().getFullYear();
  }
})();
