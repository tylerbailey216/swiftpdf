(function () {
  "use strict";

  var frame = document.getElementById("workspaceFrame");
  var connectionStatus = document.getElementById("connectionStatus");
  var logList = document.getElementById("agentLog");

  var PRESETS = {
    draft: {
      label: "Draft Copy",
      imagePageSize: "match",
      pdfPageFit: "trim",
      watermarkText: "DRAFT COPY",
      watermarkOpacity: "0.22",
      watermarkTone: "60",
      watermarkSize: "100",
      watermarkOrientation: "diagonal",
      footerText: "Internal draft - not for distribution",
      footerOpacity: "0.7",
      footerTone: "20",
      watermarkPages: "all",
      footerPages: "all"
    },
    review: {
      label: "Client Review",
      imagePageSize: "match",
      pdfPageFit: "original",
      watermarkText: "CLIENT REVIEW",
      watermarkOpacity: "0.18",
      watermarkTone: "55",
      watermarkSize: "92",
      watermarkOrientation: "diagonal",
      footerText: "Prepared for client review",
      footerOpacity: "0.6",
      footerTone: "25",
      watermarkPages: "all",
      footerPages: "all"
    },
    clean: {
      label: "Final Clean",
      imagePageSize: "match",
      pdfPageFit: "original",
      watermarkText: "",
      watermarkOpacity: "0.22",
      watermarkTone: "60",
      watermarkSize: "100",
      watermarkOrientation: "diagonal",
      footerText: "",
      footerOpacity: "0.7",
      footerTone: "13",
      watermarkPages: "none",
      footerPages: "none"
    }
  };

  function getTodayDateIso() {
    var now = new Date();
    var month = String(now.getMonth() + 1).padStart(2, "0");
    var day = String(now.getDate()).padStart(2, "0");
    return now.getFullYear() + "-" + month + "-" + day;
  }

  function addLog(message) {
    if (!logList) {
      return;
    }
    var item = document.createElement("li");
    var time = new Date().toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
    item.textContent = time + " - " + message;
    logList.prepend(item);
    while (logList.children.length > 10) {
      logList.removeChild(logList.lastChild);
    }
  }

  function setConnection(ready, message) {
    if (!connectionStatus) {
      return;
    }
    connectionStatus.classList.toggle("ready", ready);
    var label = connectionStatus.querySelector("span:last-child");
    if (label) {
      label.textContent = message;
    }
  }

  function getToolDocument() {
    if (!frame) {
      return null;
    }
    try {
      return frame.contentDocument || (frame.contentWindow ? frame.contentWindow.document : null);
    } catch (error) {
      return null;
    }
  }

  function withToolDocument(action) {
    var doc = getToolDocument();
    if (!doc || !doc.getElementById("mergeBtn")) {
      addLog("Please wait, the workspace is still loading.");
      return false;
    }
    action(doc);
    return true;
  }

  function setValue(doc, id, value) {
    var el = doc.getElementById(id);
    if (!el) {
      return false;
    }
    el.value = value;
    el.dispatchEvent(new Event("input", { bubbles: true }));
    el.dispatchEvent(new Event("change", { bubbles: true }));
    return true;
  }

  function clickById(doc, id) {
    var el = doc.getElementById(id);
    if (!el) {
      return false;
    }
    el.click();
    return true;
  }

  function applyPreset(name) {
    var preset = PRESETS[name];
    if (!preset) {
      return;
    }

    withToolDocument(function (doc) {
      setValue(doc, "imagePageSize", preset.imagePageSize);
      setValue(doc, "pdfPageFit", preset.pdfPageFit);
      setValue(doc, "watermarkText", preset.watermarkText);
      setValue(doc, "watermarkOpacity", preset.watermarkOpacity);
      setValue(doc, "watermarkTone", preset.watermarkTone);
      setValue(doc, "watermarkSize", preset.watermarkSize);
      setValue(doc, "watermarkOrientation", preset.watermarkOrientation);
      setValue(doc, "footerText", preset.footerText);
      setValue(doc, "footerOpacity", preset.footerOpacity);
      setValue(doc, "footerTone", preset.footerTone);

      clickById(doc, preset.watermarkPages === "all" ? "watermarkSelectAll" : "watermarkSelectNone");
      clickById(doc, preset.footerPages === "all" ? "footerSelectAll" : "footerSelectNone");
      addLog("Quick style set: " + preset.label + ".");
    });
  }

  function applyCustomText() {
    var watermarkText = document.getElementById("agentWatermarkText");
    var footerText = document.getElementById("agentFooterText");

    withToolDocument(function (doc) {
      if (watermarkText) {
        setValue(doc, "watermarkText", watermarkText.value.trim());
      }
      if (footerText) {
        setValue(doc, "footerText", footerText.value.trim());
      }
      addLog("Your text has been added.");
    });
  }

  function clearText() {
    var watermarkText = document.getElementById("agentWatermarkText");
    var footerText = document.getElementById("agentFooterText");
    if (watermarkText) {
      watermarkText.value = "";
    }
    if (footerText) {
      footerText.value = "";
    }
    withToolDocument(function (doc) {
      setValue(doc, "watermarkText", "");
      setValue(doc, "footerText", "");
      clickById(doc, "watermarkSelectNone");
      clickById(doc, "footerSelectNone");
      addLog("Watermark and bottom note text cleared.");
    });
  }

  function applyFillAndSign() {
    var signName = document.getElementById("agentSignName");
    var signDate = document.getElementById("agentSignDate");
    var signNote = document.getElementById("agentSignNote");
    var signPlacement = document.getElementById("agentSignPlacement");
    var signOpacity = document.getElementById("agentSignOpacity");

    withToolDocument(function (doc) {
      if (signName) {
        setValue(doc, "signName", signName.value.trim());
      }
      if (signDate) {
        setValue(doc, "signDate", signDate.value);
      }
      if (signNote) {
        setValue(doc, "signNote", signNote.value.trim());
      }
      if (signPlacement) {
        setValue(doc, "signPlacement", signPlacement.value);
      }
      if (signOpacity) {
        var percent = Number(signOpacity.value);
        var normalized = Number.isFinite(percent) ? Math.max(30, Math.min(100, percent)) / 100 : 0.95;
        setValue(doc, "signOpacity", normalized.toFixed(2));
      }
      addLog("Fill and sign details updated.");
    });
  }

  function clearFillAndSign() {
    var signName = document.getElementById("agentSignName");
    var signDate = document.getElementById("agentSignDate");
    var signNote = document.getElementById("agentSignNote");
    if (signName) {
      signName.value = "";
    }
    if (signDate) {
      signDate.value = "";
    }
    if (signNote) {
      signNote.value = "";
    }
    withToolDocument(function (doc) {
      setValue(doc, "signName", "");
      setValue(doc, "signDate", "");
      setValue(doc, "signNote", "");
      addLog("Fill and sign text cleared.");
    });
  }

  function focusSignaturePad() {
    withToolDocument(function (doc) {
      var canvas = doc.getElementById("signaturePad");
      if (!canvas) {
        addLog("Could not find the signature box right now.");
        return;
      }
      canvas.scrollIntoView({ behavior: "smooth", block: "center" });
      var previousOutline = canvas.style.outline;
      var previousOffset = canvas.style.outlineOffset;
      canvas.style.outline = "2px solid rgba(138, 215, 255, 0.9)";
      canvas.style.outlineOffset = "2px";
      window.setTimeout(function () {
        canvas.style.outline = previousOutline;
        canvas.style.outlineOffset = previousOffset;
      }, 1200);
      addLog("Signature box is now in view.");
    });
  }

  function uploadSignature() {
    withToolDocument(function (doc) {
      if (!clickById(doc, "signatureUpload")) {
        addLog("Could not open signature upload right now.");
        return;
      }
      addLog("Signature image picker opened.");
    });
  }

  function clearSignatureImage() {
    withToolDocument(function (doc) {
      if (!clickById(doc, "signatureClear")) {
        addLog("Could not find the clear signature button right now.");
        return;
      }
      addLog("Signature was cleared.");
    });
  }

  function applyDocOptions() {
    var imageSize = document.getElementById("agentImageSize");
    var pdfFit = document.getElementById("agentPdfFit");

    withToolDocument(function (doc) {
      if (imageSize) {
        setValue(doc, "imagePageSize", imageSize.value);
      }
      if (pdfFit) {
        setValue(doc, "pdfPageFit", pdfFit.value);
      }
      addLog("Page options updated.");
    });
  }

  function setFeaturePages(feature, mode) {
    var map = {
      watermark: { all: "watermarkSelectAll", none: "watermarkSelectNone" },
      footer: { all: "footerSelectAll", none: "footerSelectNone" },
      sign: { all: "signSelectAll", none: "signSelectNone" }
    };
    var target = map[feature] && map[feature][mode];
    if (!target) {
      return;
    }
    withToolDocument(function (doc) {
      clickById(doc, target);
    });
  }

  function openDownload() {
    withToolDocument(function (doc) {
      var link = doc.getElementById("downloadLink");
      if (!link || link.classList.contains("disabled") || link.getAttribute("aria-disabled") === "true") {
        addLog("Build your PDF first, then download.");
        return;
      }
      link.click();
      addLog("Download opened.");
    });
  }

  function bindButton(id, action, logMessage) {
    var button = document.getElementById(id);
    if (!button) {
      return;
    }
    button.addEventListener("click", function () {
      action();
      if (logMessage) {
        addLog(logMessage);
      }
    });
  }

  if (frame) {
    frame.addEventListener("load", function () {
      setConnection(true, "Workspace is ready");
      addLog("Workspace is ready.");
    });
  }

  document.querySelectorAll("[data-preset]").forEach(function (button) {
    button.addEventListener("click", function () {
      applyPreset(button.getAttribute("data-preset"));
    });
  });

  bindButton("openFilePicker", function () {
    withToolDocument(function (doc) {
      clickById(doc, "browseBtn");
    });
  }, "File picker opened.");

  bindButton("clearQueue", function () {
    withToolDocument(function (doc) {
      clickById(doc, "clearBtn");
    });
  }, "Removed all files from queue.");

  bindButton("applyCustomText", applyCustomText, null);
  bindButton("clearText", clearText, null);
  bindButton("applyDocOptions", applyDocOptions, null);
  bindButton("applyFillSign", applyFillAndSign, null);
  bindButton("clearFillSign", clearFillAndSign, null);
  bindButton("focusSignaturePad", focusSignaturePad, null);
  bindButton("uploadSignature", uploadSignature, null);
  bindButton("clearSignatureImage", clearSignatureImage, null);

  bindButton("wmAll", function () {
    setFeaturePages("watermark", "all");
  }, "Watermark turned on for all pages.");

  bindButton("wmNone", function () {
    setFeaturePages("watermark", "none");
  }, "Watermark turned off for all pages.");

  bindButton("ftAll", function () {
    setFeaturePages("footer", "all");
  }, "Bottom note turned on for all pages.");

  bindButton("ftNone", function () {
    setFeaturePages("footer", "none");
  }, "Bottom note turned off for all pages.");

  bindButton("sgAll", function () {
    setFeaturePages("sign", "all");
  }, "Fill and sign turned on for all pages.");

  bindButton("sgNone", function () {
    setFeaturePages("sign", "none");
  }, "Fill and sign turned off for all pages.");

  bindButton("buildPdf", function () {
    withToolDocument(function (doc) {
      clickById(doc, "mergeBtn");
    });
  }, "Started building your PDF.");

  bindButton("openDownload", openDownload, null);

  var signDateField = document.getElementById("agentSignDate");
  if (signDateField && !signDateField.value) {
    signDateField.value = getTodayDateIso();
  }

  addLog("Ready to go. Start with Add files.");
})();
