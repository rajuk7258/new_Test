document.getElementById('form8').addEventListener('submit', function(e) {
  e.preventDefault();

  const formData = {};
  new FormData(this).forEach((value, key) => formData[key] = value);

  fetch('form8_template.docx')
    .then(response => response.arrayBuffer())
    .then(template => {
      const zip = new PizZip(template);
      const doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

      doc.setData(formData);

      try {
        doc.render();
      } catch (error) {
        console.error("Error rendering template:", error);
        alert("There was an error generating the document.");
        return;
      }

      const out = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      });

      saveAs(out, "FORM-8_Report.docx");
    })
    .catch(err => {
      console.error("Error fetching template:", err);
      alert("Failed to load template file.");
    });
});
