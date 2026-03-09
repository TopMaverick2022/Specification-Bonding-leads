import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './SpecificationBondingLeads.module.scss';
import type { ISpecificationBondingLeadsProps } from './ISpecificationBondingLeadsProps';
import { sp } from "@pnp/sp/presets/all";
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';

export interface IListItem {
  Title: string;
  Description: string;
}

export default function SpecificationBondingLeads(props: ISpecificationBondingLeadsProps) {
  const [items, setItems] = useState<IListItem[]>([]);
  const [loading, setLoading] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    const fetchData = async () => {
      setLoading(true);
      try {
        const listName = "Specification Bonding leads";
        const data: any[] = await sp.web.lists.getByTitle(listName).items.select("Title", "Description").orderBy("ID", true).get();
        
        const mappedItems: IListItem[] = data.map(item => ({
          Title: item.Title || "",
          Description: item.Description || ""
        }));

        setItems(mappedItems);
      } catch (err) {
        console.error("Error fetching list data:", err);
        setError("Failed to load data. Please ensure the list 'Specification Bonding leads' exists.");
      } finally {
        setLoading(false);
      }
    };

    fetchData();
  }, []);

  const generatePdf = () => {
    const doc = new jsPDF();
    const pageHeight = doc.internal.pageSize.height;
    const pageWidth = doc.internal.pageSize.width;
    const margin = 15; // Left/Right margin
    const topMargin = 40; // Top margin to accommodate header
    const bottomMargin = 30; // Bottom margin to accommodate footer
    let y = topMargin;

    // Title Page
    doc.setFont("helvetica", "bold");
    doc.setFontSize(24);
    const titleText = "Specification Bonding leads with XLPE insulation";
    const splitTitle = doc.splitTextToSize(titleText, pageWidth - (margin * 2));
    doc.text(splitTitle, pageWidth / 2, pageHeight / 2 - 10, { align: "center" });
    
    doc.addPage();

    // Table of Contents
    y = topMargin;
    doc.setFont("helvetica", "bold");
    doc.setFontSize(16);
    doc.text("Content", margin, y);
    y += 15;

    const tocData = [
      { id: "1", title: "General", page: "2" },
      { id: "2", title: "Type Test", page: "3" },
      { id: "3", title: "Range of type approval", page: "3" },
      { id: "4", title: "Single core bonding lead", page: "3" },
      { id: "4.1", title: "Conductor", page: "3" },
      { id: "4.2", title: "Insulation", page: "3" },
      { id: "4.3", title: "Sheath", page: "3" },
      { id: "4.4", title: "Marking", page: "4" },
      { id: "4.5", title: "Routine and Sample Tests", page: "4" },
      { id: "4.6", title: "Technical data", page: "5" },
      { id: "5", title: "Concentric bonding lead", page: "6" },
      { id: "5.1", title: "Conductor", page: "6" },
      { id: "5.2", title: "Inner Insulation / Insulation", page: "6" },
      { id: "5.3", title: "Metallic Screen", page: "7" },
      { id: "5.4", title: "Outer Insulation / Sheath", page: "7" },
      { id: "5.5", title: "Marking", page: "7" },
      { id: "5.6", title: "Routine and Sample Test", page: "8" },
      { id: "5.7", title: "Technical data", page: "9" },
      { id: "", title: "Table of Modifications", page: "11" },
    ];

    tocData.forEach(row => {
      if (y + 8 > pageHeight - bottomMargin) {
        doc.addPage();
        y = topMargin;
      }
    
      doc.setFont("helvetica", "normal");
      doc.setFontSize(11);
    
      if (row.id) {
        doc.text(row.id, margin, y);
        doc.text(row.title, margin + 20, y);
      } else {
        doc.text(row.title, margin, y);
      }
    
      doc.text(row.page, pageWidth - margin, y, { align: "right" });
    
      y += 8;
    });

    doc.addPage();
    y = topMargin;

    const printSectionTitle = (title: string) => {
      if (y + 10 > pageHeight - bottomMargin) {
        doc.addPage();
        y = topMargin;
      }
    
      doc.setFont("times", "bold");   // closest to Georgia
      doc.setFontSize(12);
      doc.setTextColor(0, 0, 128);    // navy blue
    
      doc.text(title, margin, y);
    
      doc.setTextColor(0, 0, 0);      // reset
      y += 8;
    };
    
    // Helper to draw a table
    const drawTable = (title: string, headers: string[], data: string[][], colWidths: number[], columnStylesArg: Record<number, Record<string, unknown>> = {}) => {
      if (y + 15 > pageHeight - bottomMargin) { doc.addPage(); y = topMargin; }
      
      if (title) {
        printSectionTitle(title);
      }

      const columnStyles: Record<number, Record<string, unknown>> = {};
      colWidths.forEach((width, index) => {
        columnStyles[index] = { cellWidth: width };
      });

      for (const key in columnStylesArg) {
        const numKey = Number(key);
        if (columnStyles[numKey]) {
            columnStyles[numKey] = { ...columnStyles[numKey], ...columnStylesArg[numKey] };
        } else {
            columnStyles[numKey] = columnStylesArg[numKey];
        }
      }

      autoTable(doc, {
        startY: y,
        head: [headers],
        body: data,
        margin: { left: margin, right: margin, top: topMargin, bottom: bottomMargin },
        theme: "grid",
        styles: {
          font: "helvetica",
          fontSize: 10,
          cellPadding: 2,
          lineWidth: 0.2,
          overflow: "linebreak",
          valign: "middle"
        },
        headStyles: {
          fontStyle: "bold",
          fillColor: false,
          textColor: 0,
          halign: "center"
        },
        bodyStyles: {
          textColor: 0
        },
        columnStyles: columnStyles
      });

      y = (doc as any).lastAutoTable.finalY + 10;
      if (y > pageHeight - bottomMargin) {
        doc.addPage();
        y = topMargin;
      }
    };

    // Static Data for Tables
    const table45Data = [
      ["Electrical test on sheath with 25 kV DC, 1 min.", "IEC 60229", "no breakdown", "100% / 10%"],
      ["DC conductor resistance at 20°C", "IEC 60502-1 & 2, IEC 60228", "Compliance with IEC 60502-1 & IEC 60228", "100% / 10%"],
      ["Check of conductor construction", "IEC 60502-1 & 2 IEC 60228", "Compliance with IEC 60502-1 & IEC 60228", "--- / 10%"],
      ["Measurement of insulation thickness", "IEC 60811-201, 202, 203 IEC 60811-201 & IEC 60502-1 & 2", "min. average = 4,6mm", "100% / 1 sample"],
      ["Measurement of sheath thickness", "IEC 60811-201, 202, 203 IEC 60811-201 & IEC 60502-1 & 2", "min. average = 1,5 mm", "100% / 1 sample"],
      ["Hot Set Test", "IEC 60811-507 & IEC 60502-1 & 2", "elongation under load max. 175%\npermanent elongation after cooling max. 15%", "100% / 10%"],
      ["Elongation test before ageing", "", "Min. 200%", "--- / 10%"],
      ["Tensile strength test before ageing", "", "Min. 12,5 N/mm²", "--- / 10%"],
      ["Water penetration test, Annex F", "IEC 60840 Annex F & IEC 60502-1 & 2", "1 sample", "1 sample"]
    ];

    const table46Data = [
      ["Conductor", "120mm² (watertight)", "240mm² (watertight)", "300mm² (watertight)", "400mm² (watertight)", "500mm² (watertight)"],
      ["Conductor diameter", "Approx. 12,8mm", "Approx. 18mm", "Approx. 20mm", "Approx. 23mm", "Approx. 26mm"],
      ["Conductor construction", "Round, stranded", "Round, stranded", "Round, stranded", "Round, stranded", "Round, stranded"],
      ["Insulation (min. avg.)", "4,6mm", "4,6mm", "4,6mm", "4,6mm", "4,6mm"],
      ["Sheath (min. avg.)", "1,5mm", "1,5mm", "1,5mm", "1,5mm", "1,5mm"],
      ["Extruded semiconductive layer", "Yes", "Yes", "Yes", "Yes", "Yes"],
      ["Outer diameter (OD)", "Approx. 27mm", "Approx. 35mm", "Approx. 38mm", "Approx. 39mm", "Approx. 41mm"],
      ["Rated Voltage in kV", "6/10", "6/10", "6/10", "6/10", "6/10"],
      ["Test Voltage on insulation + sheath", "25kV", "25kV", "25kV", "25kV", "25kV"],
      ["Operating conductor temperature", "90°C", "90°C", "90°C", "90°C", "90°C"],
      ["Max. short circuit temperature", "250°C", "250°C", "250°C", "250°C", "250°C"],
      ["Operating temperature range", "-35 to 90°C", "-35 to 90°C", "-35 to 90°C", "-35 to 90°C", "-35 to 90°C"],
      ["Min. temperature for laying", "0°C", "0°C", "0°C", "0°C", "0°C"],
      ["Min. storage temperature", "-35°C", "-35°C", "-35°C", "-35°C", "-35°C"],
      ["Colour of insulation", "Black", "Black", "Black", "Black", "Black"],
      ["Colour of outer sheath", "Nature or grey", "Nature or grey", "Nature or grey", "Nature or grey", "Nature or grey"],
      ["Colour of extruded semiconductive outer sheath", "Black", "Black", "Black", "Black", "Black"],
      ["UV stability", "Yes", "Yes", "Yes", "Yes", "Yes"],
      ["Max. DC resistance @20°C in W/km", "0,1530", "0,0754", "0,0601", "0,047", "0,0366"],
      ["Short circuit current (kA/1sec)", "20,4", "40,9", "51,1", "68,1", "85,2"],
      ["Min. Bending radius", "15xOD", "15xOD", "15xOD", "15xOD", "15xOD"],
      ["Suitable for laying in", "Air & soil", "Air & soil", "Air & soil", "Air & soil", "Air & soil"]
    ];

    const table56Data = [
      ["DC conductor resistance at 20°C", "IEC 60502-1 & 2 IEC 60228", "Compliance with IEC 60502-1 & IEC 60228", "100% / 10%"],
      ["Electrical test on cable inner insulation (XLPE)", "IEC 60229", "no breakdown", "100% / 10%"],
      ["Electrical test on cable outer insulation / sheath", "IEC 60229", "no breakdown", "100% / 10%"],
      ["Check of conductor construction", "IEC 60502-1 & 2 IEC 60228", "Compliance with IEC 60502-1 & IEC 60228", "--- / 10%"],
      ["Measuring of inner XLPE insulation thickness", "IEC 60811...", "min. average = 4,6mm / 7,0mm", "100% / 1 sample"],
      ["Measuring of outer insulation / sheath thickness", "IEC 60811...", "min. average = 3,3 mm", "100% / 1 sample"],
      ["Hot Set Test", "IEC 60811...", "elongation max 175%", "100% / 10%"],
      ["Elongation test before ageing", "", "Min. 200%", "--- / 10%"],
      ["Tensile strength test before ageing", "", "Min. 12,5 N/mm²", "--- / 10%"],
      ["Water penetration test, Annex E", "IEC 60840 Annex E", "1 sample", "1 sample"],
      ["Water penetration test, Annex F", "IEC 60840 Annex F", "1 sample", "1 sample"]
    ];

    const table57Data = [
      ["Conductor", "120mm²", "240mm²", "300mm²", "400mm²", "500mm²"],
      ["Conductor diameter", "12,8mm", "18mm", "20mm", "23mm", "26mm"],
      ["Conductor construction", "Round, stranded", "Round, stranded", "Round, stranded", "Round, stranded", "Round, stranded"],
      ["Inner Insulation (min. avg.)", "4,6mm", "4,6mm", "4,6mm", "7,0mm", "7,0mm"],
      ["Metallic screen", "120mm²", "240mm²", "300mm²", "400mm²", "500mm²"],
      ["Outer Insulation / sheath (min.)", "3,3mm", "3,3mm", "3,3mm", "3,3mm", "3,3mm"],
      ["extruded semiconductive layer", "Yes", "Yes", "Yes", "Yes", "Yes"],
      ["Outer diameter (OD)", "36mm", "44mm", "51mm", "54mm", "62mm"],
      ["Rated Voltage in kV", "6/10", "6/10", "6/10", "6/10", "6/10"],
      ["Test Voltage inner (4,6mm)", "15kV AC", "15kV AC", "15kV AC", "15kV AC", "15kV AC"],
      ["Test Voltage inner (7,0mm)", "20kV AC", "20kV AC", "20kV AC", "20kV AC", "20kV AC"],
      ["Test Voltage outer", "25kV DC", "25kV DC", "25kV DC", "25kV DC", "25kV DC"],
      ["Operating conductor temp", "90°C", "90°C", "90°C", "90°C", "90°C"],
      ["Max. short circuit temp", "250°C", "250°C", "250°C", "250°C", "250°C"],
      ["Operating temp range", "-35 to 90°C", "-35 to 90°C", "-35 to 90°C", "-35 to 90°C", "-35 to 90°C"],
      ["Min. temp for laying", "0°C", "0°C", "0°C", "0°C", "0°C"],
      ["Min. storage temp", "-35°C", "-35°C", "-35°C", "-35°C", "-35°C"],
      ["Colour inner", "Black", "Black", "Black", "Black", "Black"],
      ["Colour outer", "Nature/grey", "Nature/grey", "Nature/grey", "Nature/grey", "Nature/grey"],
      ["Colour semi", "Black", "Black", "Black", "Black", "Black"],
      ["UV stability", "Yes", "Yes", "Yes", "Yes", "Yes"],
      ["Min. DC resistance", "0,153", "0,0754", "0,0601", "0,047", "0,0366"],
      ["Short circuit current", "20,4", "40,9", "51,1", "68,1", "85,2"],
      ["Min. Bending radius", "15xOD", "15xOD", "15xOD", "15xOD", "15xOD"],
      ["Suitable for laying in", "Air & soil", "Air & soil", "Air & soil", "Air & soil", "Air & soil"]
    ];

    const tableModifications = [
      ["A", "21.12.2023", "Oliver Sablic", "First issue"],
      ["B", "04.01.2023", "Oliver Sablic", "graphite coating deleted"],
      ["C", "05.03.2024", "Oliver Sablic", "Adjusting of short circuit current"],
      ["D", "09.04.2024", "Oliver Sablic", "Added 145kV lighting impulse test, adjusting meter marking"],
      ["E", "10.07.2024", "Oliver Sablic", "outer sheath incl. semicon-layer must be fully bonded"],
      ["F", "10.04.2025", "Oliver Sablic", "2: cables acc. to ENA Rec\n4.1 & 5.1: Tape over conductor\n5.3: Cu wire screen double layer\nOuter sheath marking, Frequency"]
    ];

    const routineTestColumnStyles = { 0: { halign: "left" }, 1: { halign: "center" }, 2: { halign: "left" }, 3: { halign: "center" } };
    const techDataColumnStyles = { 0: { halign: "left" }, 1: { halign: "center" }, 2: { halign: "center" }, 3: { halign: "center" }, 4: { halign: "center" }, 5: { halign: "center" } };

    const printSection = (title: string, description: string) => {
      if (y + 10 > pageHeight - bottomMargin) { doc.addPage(); y = topMargin; }
      
      doc.setFont("times", "bold");
      doc.setTextColor(0, 0, 128);

      if (/^\d+\.\d+/.test(title)) {
        doc.setFontSize(12);
      } else {
        doc.setFontSize(14.5);
      }

      doc.text(title, margin, y);
      y += 8;

      if (description) {
        const lines = description.split(/\r?\n/);

        lines.forEach(line => {
          const trimmedLine = line.trim();
          if (!trimmedLine) {
            y += 2;
            return;
          }

          const isSubtitle = /^\d+\.\d+/.test(trimmedLine);

          if (isSubtitle) {
            doc.setFont("times", "bold");
            doc.setFontSize(12);
            doc.setTextColor(0, 0, 128);
          } else {
            doc.setFont("helvetica", "normal");
            doc.setFontSize(11);
            doc.setTextColor(0, 0, 0);
          }

          const splitText = doc.splitTextToSize(trimmedLine, pageWidth - (margin * 2));
          const lineHeight = 6;
          const textHeight = splitText.length * lineHeight;

          if (y + textHeight > pageHeight - bottomMargin) {
            doc.addPage();
            y = topMargin;
          }

          doc.text(splitText, margin, y);
          y += textHeight + 2;
        });
      }
      y += 5;
    };

    // Iterate through all items and print them
    items.forEach(item => {
      printSection(item.Title, item.Description);
      
      const title = item.Title.toLowerCase().trim();
      
      if (title.indexOf("single core bonding lead") !== -1) {
        drawTable("4.5 Routine and Sample Tests", ["Test procedure", "Standard", "Acceptance", "Freq (Routine/Sample)"], table45Data, [70, 35, 55, 30], routineTestColumnStyles);
        drawTable("4.6 Technical data", ["Property", "Type 120", "Type 240", "Type 300", "Type 400", "Type 500"], table46Data, [40, 26, 26, 26, 26, 26], techDataColumnStyles);
      } else if (title.indexOf("concentric bonding lead") !== -1) {
        drawTable("5.6 Routine and Sample Test", ["Test procedure", "Standard", "Acceptance", "Freq (Routine/Sample)"], table56Data, [50, 40, 50, 30], routineTestColumnStyles);
        drawTable("5.7 Technical data", ["Property", "Type 120/120", "Type 240/240", "Type 300/300", "Type 400/400", "Type 500/500"], table57Data, [40, 26, 26, 26, 26, 26], techDataColumnStyles);
      }
    });

    // Table of Modifications at the end
    doc.addPage();
    y = topMargin;
    drawTable("Table of Modifications", ["Rev.", "Date", "Prepared by", "Description"], tableModifications, [15, 25, 30, 100]);

    // Add Header and Footer to all pages
    const totalPages = doc.getNumberOfPages();
    for (let i = 2; i <= totalPages; i++) {
      doc.setPage(i);
      
      // --- Header ---
      const headerY = 10;
      const headerColWidth = (pageWidth - (margin * 2)) / 3;
      
      doc.setDrawColor(0);
      doc.setLineWidth(0.1);
      doc.setTextColor(0);

      // Header Line
      doc.line(margin, headerY + 12, pageWidth - margin, headerY + 12);

      // Row 1
      doc.setFont("helvetica", "normal");
      
      doc.setFontSize(6);
      doc.text("Doc. ID.:", margin, headerY + 4);
      doc.setFontSize(8);
      doc.text("1AA0648051", margin + 25, headerY + 4);

      doc.setFontSize(6);
      doc.text("Classification:", margin + headerColWidth, headerY + 4);
      doc.setFontSize(8);
      doc.text("Technical specification", margin + headerColWidth + 25, headerY + 4);

      doc.setFontSize(6);
      doc.text("Prepared by:", margin + (headerColWidth * 2), headerY + 4);
      doc.setFontSize(8);
      doc.text("Oliver Sablic", margin + (headerColWidth * 2) + 25, headerY + 4);

      // Row 2
      doc.setFontSize(6);
      doc.text("Revision:", margin, headerY + 9);
      doc.setFontSize(8);
      doc.text("F", margin + 25, headerY + 9);

      doc.setFontSize(6);
      doc.text("Project ID:", margin + headerColWidth, headerY + 9);
      doc.setFontSize(8);
      doc.text("BerLEAN", margin + headerColWidth + 25, headerY + 9);

      doc.setFontSize(6);
      doc.text("Approved by:", margin + (headerColWidth * 2), headerY + 9);
      doc.setFontSize(8);
      doc.text("Andre Sobolewski-Dockal", margin + (headerColWidth * 2) + 25, headerY + 9);

      // --- Footer ---
      const footerY = pageHeight - 15;
      
      // Footer Line
      doc.line(margin, footerY - 5, pageWidth - margin, footerY - 5);
      
      // Footer Text
      doc.setFont("helvetica", "normal");
      doc.setFontSize(8);
      doc.setTextColor(0);
      doc.text("Copyright 2026 NKT GmbH & Co. KG. All rights reserved.", margin, footerY);
      
      doc.setFont("times", "bold");
      doc.setFontSize(9);
      doc.setTextColor(0, 0, 128);
      doc.text(`Page ${i} / ${totalPages}`, pageWidth - margin, footerY, { align: "right" });
    }

    doc.save("Specification_Bonding_Leads.pdf");
  };

  return (
    <section className={`${styles.specificationBondingLeads} ${props.hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
             
        {loading && <p>Loading data...</p>}
        {error && <p style={{ color: 'red' }}>{error}</p>}
        
        {!loading && !error && (
          <button className={styles.pdfButton} onClick={generatePdf}>
            Download PDF
          </button>
        )}
      </div>
    </section>
  );
}