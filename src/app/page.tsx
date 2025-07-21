"use client";

import { useState } from "react";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import { saveAs } from "file-saver";
import {
  Box,
  Button,
  Container,
  TextField,
  Typography,
  Paper,
  List,
  ListItem,
  Link,
  CssBaseline,
} from "@mui/material";

export default function Home() {
  const [fileName, setFileName] = useState<string>("");
  const [splitFiles, setSplitFiles] = useState<{ name: string; blob: Blob }[]>([]);
  const [numParts, setNumParts] = useState<number>(2);

  const handleFile = async (file: File) => {
    setFileName(file.name);
    const data = await file.arrayBuffer();
    const isCSV = file.name.endsWith(".csv");
    const workbook = isCSV
      ? XLSX.read(data, { type: "array", codepage: 65001 })
      : XLSX.read(data, { type: "array" });

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData: any[] = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const header = jsonData[0];
    const rows = jsonData.slice(1);
    const chunkSize = Math.ceil(rows.length / numParts);
    const chunks = Array.from({ length: numParts }, (_, i) =>
      [header, ...rows.slice(i * chunkSize, (i + 1) * chunkSize)]
    );

    const baseName = file.name.replace(/\.(xlsx|csv)$/, "");
    const ext = isCSV ? "csv" : "xlsx";

    const newFiles = chunks.map((chunk, i) => {
      const partNumber = (i + 1).toString().padStart(2, "0");
      const name = `${baseName}_Part${partNumber}.${ext}`;
      const ws = XLSX.utils.aoa_to_sheet(chunk);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
      const blob = new Blob(
        [
          XLSX.write(wb, {
            bookType: ext as any,
            type: "array",
          }),
        ],
        { type: "application/octet-stream" }
      );
      return { name, blob };
    });

    setSplitFiles(newFiles);
  };

  const handleDrop = (e: React.DragEvent<HTMLDivElement>) => {
    e.preventDefault();
    const file = e.dataTransfer.files[0];
    if (file && (file.name.endsWith(".xlsx") || file.name.endsWith(".csv"))) {
      handleFile(file);
    }
  };

  const handleDownloadAll = async () => {
    const zip = new JSZip();
    splitFiles.forEach(file => {
      zip.file(file.name, file.blob);
    });
    const content = await zip.generateAsync({ type: "blob" });
    saveAs(content, `${fileName.replace(/\.(xlsx|csv)/, "")}_split.zip`);
  };

  return (
    <Box component="main" sx={{ bgcolor: "#fafafa", minHeight: "100vh", py: 4 }}>
      <CssBaseline />
      <Container maxWidth="sm">
        <Typography variant="h4" component="h1" gutterBottom align="center" color="primary">
          Excel / CSV File Splitter
        </Typography>

        <TextField
          label="Number of parts"
          type="number"
          fullWidth
          variant="outlined"
          margin="normal"
          value={numParts}
          onChange={(e) => setNumParts(Number(e.target.value))}
        />

        <Paper
          variant="outlined"
          onDrop={handleDrop}
          onDragOver={(e) => e.preventDefault()}
          sx={{ p: 4, mt: 2, textAlign: "center", color: "text.secondary", cursor: "pointer" }}
        >
          Drag & drop your .xlsx or .csv file here
        </Paper>

        {splitFiles.length > 0 && (
          <Box mt={4}>
            <Button
              onClick={handleDownloadAll}
              variant="contained"
              color="primary"
              fullWidth
              sx={{ mb: 2 }}
            >
              ⬇ Download All as ZIP
            </Button>

            <Paper variant="outlined" sx={{ p: 2 }}>
              <Typography variant="subtitle1">Individual Downloads</Typography>
              <List>
                {splitFiles.map((file, i) => (
                  <ListItem key={i}>
                    <Link
                      href={URL.createObjectURL(file.blob)}
                      download={file.name}
                      underline="hover"
                      color="primary"
                    >
                      {file.name}
                    </Link>
                  </ListItem>
                ))}
              </List>
            </Paper>
          </Box>
        )}
      </Container>
    </Box>
  );
}