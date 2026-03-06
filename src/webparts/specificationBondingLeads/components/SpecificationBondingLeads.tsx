import * as React from 'react';
import { useState, useEffect } from 'react';
import styles from './SpecificationBondingLeads.module.scss';
import type { ISpecificationBondingLeadsProps } from './ISpecificationBondingLeadsProps';
import { sp } from "@pnp/sp/presets/all";
import jsPDF from 'jspdf';

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
        // Fetch data using PnP JS
        const data: any[] = await sp.web.lists.getByTitle(listName).items.select("Title", "Description").get();
        
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
    const margin = 20;
    let y = 20;

    doc.setFont("helvetica", "normal");
    doc.setFontSize(11);

    const printSection = (title: string, description: string) => {
      // Check space for Title
      if (y + 10 > pageHeight - margin) { doc.addPage(); y = margin; }
      
      // Print Title
      doc.setFont("helvetica", "bold");
      doc.text(title, margin, y);
      y += 7;

      // Print Description
      doc.setFont("helvetica", "normal");
      const splitDesc = doc.splitTextToSize(description, pageWidth - (margin * 2));
      const lineHeight = 6;
      const descHeight = splitDesc.length * lineHeight;

      // Check space for Description
      if (y + descHeight > pageHeight - margin) { 
        doc.addPage(); 
        y = margin; 
      }

      doc.text(splitDesc, margin, y);
      y += descHeight + 10; // Add spacing after section
    };

    // Iterate through all items and print them
    items.forEach(item => {
      printSection(item.Title, item.Description);
    });

    doc.save("Specification_Bonding_Leads.pdf");
  };

  return (
    <section className={`${styles.specificationBondingLeads} ${props.hasTeamsContext ? styles.teams : ''}`}>
      <div className={styles.welcome}>
        <h2>Specification Bonding Leads</h2>
        
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
