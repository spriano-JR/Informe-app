class ExportService {
    constructor() {
        this.currentReportHtml = '';
    }

    async exportPDF() {
        const { jsPDF } = window.jspdf;
        const pageSize = document.getElementById('page-size').value;
        const format = pageSize === 'letter' ? 'letter' : (pageSize === 'legal' ? 'legal' : 'a4');
        const element = document.getElementById('report-paper');

        // Prepare element for full capture
        const originalStyle = element.getAttribute('style') || '';
        const originalHeight = element.style.height;
        const originalOverflow = element.style.overflow;

        // Expand element to its full content height
        element.style.height = 'auto';
        element.style.overflow = 'visible';

        try {
            const canvas = await html2canvas(element, {
                scale: 2,
                useCORS: true,
                logging: false,
                backgroundColor: '#ffffff',
                windowWidth: element.scrollWidth,
                windowHeight: element.scrollHeight
            });

            const imgData = canvas.toDataURL('image/png');
            const doc = new jsPDF('p', 'mm', format);

            const imgProps = doc.getImageProperties(imgData);
            const pdfWidth = doc.internal.pageSize.getWidth();
            const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;

            let heightLeft = pdfHeight;
            let position = 0;
            const pageHeight = doc.internal.pageSize.getHeight();

            doc.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
            heightLeft -= pageHeight;

            while (heightLeft >= 0) {
                position = heightLeft - pdfHeight;
                doc.addPage();
                doc.addImage(imgData, 'PNG', 0, position, pdfWidth, pdfHeight);
                heightLeft -= pageHeight;
            }

            doc.save('Report-Master-Pro.pdf');
        } catch (error) {
            console.error('Error generating PDF:', error);
            alert('Error al generar PDF: ' + error.message);
        } finally {
            // Restore original styles
            element.setAttribute('style', originalStyle);
            element.style.height = originalHeight;
            element.style.overflow = originalOverflow;
        }
    }

    async exportWord() {
        try {
            const docxLib = window.docx || window.Docx;
            if (!docxLib) {
                throw new Error("La librería de Word (docx) no se encontró. Intenta presionar (Ctrl + F5).");
            }

            const pageSizeSelect = document.getElementById('page-size');
            const pageSize = pageSizeSelect ? pageSizeSelect.value : 'a4';
            const width = pageSize === 'letter' ? 12240 : (pageSize === 'legal' ? 12240 : 11906);
            const height = pageSize === 'letter' ? 15840 : (pageSize === 'legal' ? 20160 : 16838);

            const logoImg = document.querySelector('#paper-logo img');
            const signImg = document.querySelector('#paper-signature img');
            const reportTitleEl = document.getElementById('report-title');
            const reportBodyEl = document.getElementById('report-body');
            const footerParaEl = document.querySelector('#report-footer p.font-bold');

            const reportTitle = reportTitleEl ? reportTitleEl.textContent : "INFORME";
            const footerText = footerParaEl ? footerParaEl.textContent : "Consultoría Estratégica";

            const children = [];
            const isDataUrl = (str) => str && str.startsWith('data:image');

            // 1. Add Logo
            if (logoImg && isDataUrl(logoImg.src)) {
                try {
                    const logoBase64 = logoImg.src.split(',')[1];
                    children.push(new docxLib.Paragraph({
                        children: [
                            new docxLib.ImageRun({
                                data: Uint8Array.from(atob(logoBase64), c => c.charCodeAt(0)),
                                transformation: { width: 60, height: 60 },
                            }),
                        ],
                    }));
                } catch (e) { console.error("Logo export error:", e); }
            }

            // 2. Add Title
            children.push(new docxLib.Paragraph({
                alignment: docxLib.AlignmentType.CENTER,
                children: [
                    new docxLib.TextRun({
                        text: reportTitle,
                        bold: true,
                        size: 32,
                        color: "05070A",
                    }),
                ],
                spacing: { before: 400, after: 600 },
            }));

            // 3. Process Body Content (Text + Charts + Insights)
            const bodyNodes = reportBodyEl.childNodes;
            for (const node of bodyNodes) {
                // Handle Insight Cards Grid
                if (node.nodeType === 1 && node.classList.contains('grid')) {
                    const cards = node.querySelectorAll('.insight-card');
                    cards.forEach(card => {
                        const type = card.querySelector('.badge-critical') ? "HALLAZGO CRÍTICO" : "OPORTUNIDAD";
                        const content = card.querySelector('p').innerText;

                        children.push(new docxLib.Paragraph({
                            shading: { fill: card.querySelector('.badge-critical') ? "FFEEEE" : "EEFFEE", type: docxLib.ShadingType.CLEAR, color: "auto" },
                            border: { left: { color: card.querySelector('.badge-critical') ? "FF0000" : "00FF00", space: 1, style: docxLib.BorderStyle.SINGLE, size: 12 } },
                            children: [
                                new docxLib.TextRun({ text: `${type}: `, bold: true, color: card.querySelector('.badge-critical') ? "BB0000" : "008800" }),
                                new docxLib.TextRun({ text: content }),
                            ],
                            spacing: { before: 200, after: 200 },
                        }));
                    });
                }

                // Handle Chart Containers
                else if (node.nodeType === 1 && node.classList.contains('chart-container')) {
                    const canvas = node.querySelector('canvas');
                    const title = node.querySelector('p').innerText;
                    if (canvas) {
                        try {
                            const chartDataUrl = canvas.toDataURL('image/png', 1.0);
                            const chartBase64 = chartDataUrl.split(',')[1];

                            children.push(new docxLib.Paragraph({
                                text: title,
                                bold: true,
                                spacing: { before: 400, after: 200 },
                            }));

                            children.push(new docxLib.Paragraph({
                                alignment: docxLib.AlignmentType.CENTER,
                                children: [
                                    new docxLib.ImageRun({
                                        data: Uint8Array.from(atob(chartBase64), c => c.charCodeAt(0)),
                                        transformation: { width: 500, height: 250 },
                                    }),
                                ],
                                spacing: { after: 400 },
                            }));
                        } catch (e) { console.error("Chart export error:", e); }
                    }
                }

                // Handle Standard Prose
                else if (node.nodeType === 1 && node.classList.contains('prose')) {
                    const proseContent = node.innerHTML.split('<br>');
                    proseContent.forEach(line => {
                        const cleanLine = line.replace(/<[^>]*>/g, '').trim();
                        if (cleanLine) {
                            children.push(new docxLib.Paragraph({
                                text: cleanLine,
                                spacing: { after: 150 },
                            }));
                        }
                    });
                }
            }

            // 4. Add Signature
            if (signImg && isDataUrl(signImg.src)) {
                try {
                    const signBase64 = signImg.src.split(',')[1];
                    children.push(new docxLib.Paragraph({
                        alignment: docxLib.AlignmentType.CENTER,
                        spacing: { before: 800 },
                        children: [
                            new docxLib.ImageRun({
                                data: Uint8Array.from(atob(signBase64), c => c.charCodeAt(0)),
                                transformation: { width: 120, height: 60 },
                            }),
                        ],
                    }));
                } catch (e) { console.error("Signature export error:", e); }
            }

            // 5. Add Footer Text
            children.push(new docxLib.Paragraph({
                alignment: docxLib.AlignmentType.CENTER,
                children: [
                    new docxLib.TextRun({
                        text: footerText,
                        bold: true,
                    }),
                ],
            }));

            const doc = new docxLib.Document({
                sections: [{
                    properties: {
                        page: { size: { width, height } },
                    },
                    children: children,
                }],
            });

            const blob = await docxLib.Packer.toBlob(doc);
            saveAs(blob, "Report-Master-Pro.docx");
        } catch (error) {
            console.error('Error generating Word:', error);
            alert('Error al generar Word: ' + error.message);
        }
    }

    async exportExcel(data) {
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Datos Analizados");
        XLSX.writeFile(wb, "Report-Master-Pro.xlsx");
    }

    async exportPPT() {
        let pres = new PptxGenJS();
        let slide = pres.addSlide();

        const logoImg = document.querySelector('#paper-logo img');
        if (logoImg && logoImg.src) {
            slide.addImage({ data: logoImg.src, x: 0.5, y: 0.5, w: 0.8, h: 0.8 });
        }

        slide.addText("REPORT-MASTER PRO", { x: 1.5, y: 0.7, fontSize: 24, bold: true, color: '00949D' });
        slide.addText(document.getElementById('report-body').innerText.substring(0, 1000), { x: 1, y: 2, w: 8, h: 4, fontSize: 12 });

        pres.writeFile({ fileName: "Report-Master-Pro.pptx" });
    }
}

// Utility to save blob (since FileSaver is often needed for docx)
function saveAs(blob, fileName) {
    const link = document.createElement('a');
    link.href = window.URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
}

const exportService = new ExportService();
