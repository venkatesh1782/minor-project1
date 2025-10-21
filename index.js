// ---------------------------
// Template selection logic
// ---------------------------
function selectTemplate(templateId) {
  const templates = document.querySelectorAll(".template-thumb");
  templates.forEach((t) => t.classList.remove("selected"));
  const selected = document.getElementById(templateId);
  selected.classList.add("selected");
  document.getElementById("template-input").value = templateId;

  const previewContainer = document.getElementById("template-preview");
  previewContainer.innerHTML = `
    <h3>Selected Template Preview:</h3>
    <img src="${selected.src}" class="preview-image" alt="Selected Template">
  `;
}

// ---------------------------
// Utility Functions
// ---------------------------
function updateImageCount() {
  const input = document.getElementById("images-input");
  document.getElementById("total-images").textContent =
    "Total Images: " + input.files.length;
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => resolve(e.target.result);
    reader.onerror = (e) => reject(e);
    reader.readAsDataURL(file);
  });
}

async function imgToBase64(url) {
  const response = await fetch(url);
  const blob = await response.blob();
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

// ---------------------------
// Generate PPT Logic
// ---------------------------
document.querySelector(".generate-btn").addEventListener("click", async function (e) {
  e.preventDefault();
  const loadingLine = document.getElementById("loading-line");
  if (loadingLine) {
    loadingLine.style.display = "block";
    loadingLine.textContent = "⏳ Generating PPT, please wait...";
    loadingLine.style.color = "#0078D4";
  }

  try {
    const fields = [
      "study_from", "display_title", "project_owner",
      "prepared_by", "submitted_on", "client",
      "doc_no", "project", "surveyed_on", "surveyed_by"
    ];
    const formData = {};
    fields.forEach(f => formData[f] = document.getElementById(f)?.value || "");

    const templateId = document.getElementById("template-input").value;
    const templateImg = document.getElementById(templateId);
    if (!templateImg) {
      alert("Please select a PPT template first!");
      if (loadingLine) loadingLine.style.display = "none";
      return;
    }

    const imageFiles = document.getElementById("images-input").files;
    if (!imageFiles.length) {
      alert("Please select images to include in the PPT.");
      if (loadingLine) loadingLine.style.display = "none";
      return;
    }

    const pptx = new PptxGenJS();

    // ---------------------------
    // Template-based backgrounds
    // ---------------------------
    let frontImagePath = "";
    let lastImagePath = "";

    if (templateId === "template1" || templateId === "template3") {
      frontImagePath = "ppt1 template1.jpg";
      lastImagePath = "ppt1 last page.jpg";
    } else if (templateId === "template2") {
      frontImagePath = "ppt2 template2.jpg";
      lastImagePath = "ppt2 lastpage.jpg";
    }

    const templateBase64 = await imgToBase64(templateImg.src);
    const frontBase64 = await imgToBase64(frontImagePath);
    const lastBase64 = await imgToBase64(lastImagePath);

    pptx.defineSlideMaster({
      title: "TemplateMaster",
      background: { data: templateBase64 },
    });

    // ---------------------------
    // Slide 1: Front Cover
    // ---------------------------
    const frontSlide = pptx.addSlide();
    frontSlide.addImage({ data: frontBase64, x: 0, y: 0, w: 10, h: 7.5 });

    // ---------------------------
    // Slide 2: Table Details
    // ---------------------------
    const tableSlide = pptx.addSlide({ masterName: "TemplateMaster" });
    const tableData = [
      [
        {
          text: formData.display_title || "Project Information",
          options: { colspan: 4, bold: true, fontSize: 20, align: "center", fill: "F2F2F2" },
        },
      ],
      [
        { text: "Study From:", options: { bold: true } }, formData.study_from,
        { text: "Project:", options: { bold: true } }, formData.project,
      ],
      [
        { text: "Client:", options: { bold: true } }, formData.client,
        { text: "Project Owner:", options: { bold: true } }, formData.project_owner,
      ],
      [
        { text: "Doc No:", options: { bold: true } }, formData.doc_no,
        { text: "Surveyed On:", options: { bold: true } }, formData.surveyed_on,
      ],
      [
        { text: "Surveyed By:", options: { bold: true } }, formData.surveyed_by,
        { text: "Prepared By:", options: { bold: true } }, formData.prepared_by,
      ],
      [
        { text: "Submitted On:", options: { bold: true } }, formData.submitted_on, "", "",
      ],
    ];

    tableSlide.addTable(tableData, {
      x: 0.6, y: 1.0, w: 9,
      colW: [1.5, 2.3, 1.7, 3.0],
      border: { pt: 1, color: "000000" },
      fontSize: 14, color: "000000", align: "left",
    });

    tableSlide.addText("NTC Logistics India Pvt Limited", {
      x: 0, y: 5.8, w: 10, fontSize: 14, bold: true, color: "000000", align: "center",
    });

    // ---------------------------
    // Slides 3+: Images (Auto-fit with template-specific padding)
    // ---------------------------
     // Slides 3+: Images (auto-fit ~65%)
    // ---------------------------
    for (let file of imageFiles) {
      const imgBase64 = await fileToBase64(file);
      const slide = pptx.addSlide({ masterName: "TemplateMaster" });

      await new Promise((resolve) => {
        const img = new Image();
        img.src = imgBase64;
        img.onload = () => {
          const slideW = 10.0;
          const slideH = 7.5;
          const padX = 0.4;
          const padY = 0.4;
          const maxW = slideW - padX * 2;
          const maxH = (slideH - padY * 2) * 0.65;

          const imgRatio = img.width / img.height;
          const areaRatio = maxW / maxH;
          let w, h;

          if (imgRatio > areaRatio) {
            w = maxW;
            h = w / imgRatio;
          } else {
            h = maxH;
            w = h * imgRatio;
          }

          const x = (slideW - w) / 2;
          const y = padY + (maxH - h) / 2;

          slide.addImage({
            data: imgBase64,
            x,
            y,
            w,
            h,
            sizing: { type: "contain", w, h },
          });


          resolve();
        };
      });
    }



    // ---------------------------
    // Final Slide: End Page
    // ---------------------------
    const lastSlide = pptx.addSlide();
    lastSlide.addImage({ data: lastBase64, x: 0, y: 0, w: 10, h: 7.5 });

    // ---------------------------
    // Save PPT
    // ---------------------------
    await pptx.writeFile({
      fileName: `${formData.display_title || "NTC_Presentation"}.pptx`,
    });

    if (loadingLine) {
      loadingLine.textContent = "✅ PPT generated successfully!";
      loadingLine.style.color = "green";
      setTimeout(() => (loadingLine.style.display = "none"), 3000);
    }

  } catch (err) {
    console.error(err);
    if (loadingLine) {
      loadingLine.textContent = "❌ Error generating PPT!";
      loadingLine.style.color = "red";
      setTimeout(() => (loadingLine.style.display = "none"), 4000);
    }
  }
});

// ---------------------------
// Bind updateImageCount
// ---------------------------
let imageCountTimeout;
document.getElementById("images-input").addEventListener("change", () => {
  clearTimeout(imageCountTimeout);
  imageCountTimeout = setTimeout(updateImageCount, 300);
});
