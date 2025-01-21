import React, { useState } from "react";
import {
  Button,
  Container,
  Paper,
  Typography,
  List,
  ListItem,
  ListItemText,
  CircularProgress,
  styled,
} from "@mui/material";
import CloudUploadIcon from "@mui/icons-material/CloudUpload";
import FileDownloadIcon from "@mui/icons-material/FileDownload";
import ExcelJS from "exceljs";

// Styled components
const GradientBackground = styled("div")({
  minHeight: "100vh",
  height: "100vh",
  background: "linear-gradient(180deg, #46FE36 0%, #F3F9F4 100%)",
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  fontFamily: "Montserrat, sans-serif",
});

const StyledPaper = styled(Paper)(({ theme }) => ({
  padding: "2rem",
  borderRadius: "16px",
  boxShadow: "0 8px 32px rgba(0, 0, 0, 0.1)",
  width: "100%",
  maxWidth: "800px",
  maxHeight: "90vh",
  overflowY: "auto",
}));

const BannerImage = styled("img")({
  width: "100%",
  height: "auto",
  borderRadius: "8px",
  marginBottom: "2rem",
  display: "block",
});

const UploadButton = styled(Button)({
  borderRadius: "8px",
  backgroundColor: "#000",
  color: "#fff",
  padding: "12px 24px",
  "&:hover": {
    backgroundColor: "#333",
  },
});

const GenerateButton = styled(Button)({
  border: "2px solid #000",
  color: "#000",
  padding: "10px 24px",
  backgroundColor: "transparent",
  "&:hover": {
    backgroundColor: "rgba(0, 0, 0, 0.05)",
    borderColor: "#ccc",
  },
  "&.Mui-disabled": {
    borderColor: "#ccc",
    color: "#ccc",
  },
});

const FileInput = styled("div")({
  border: "1px solid #ddd",
  borderRadius: "8px",
  padding: "1.5rem",
  marginBottom: "1.5rem",
  backgroundColor: "#fff",
});

function App() {
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [isProcessing, setIsProcessing] = useState(false);

  const handleFileSelect = (event) => {
    const files = Array.from(event.target.files);
    const imageFiles = files.filter((file) => file.type.startsWith("image/"));
    setSelectedFiles(imageFiles);
  };

  const createThumbnail = (file) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
          const canvas = document.createElement("canvas");
          const ctx = canvas.getContext("2d");

          canvas.width = 50;
          canvas.height = 50;

          const scale = Math.min(50 / img.width, 50 / img.height);
          const width = img.width * scale;
          const height = img.height * scale;

          const x = (50 - width) / 2;
          const y = (50 - height) / 2;

          ctx.fillStyle = "white";
          ctx.fillRect(0, 0, 50, 50);
          ctx.drawImage(img, x, y, width, height);

          canvas.toBlob((blob) => {
            const reader = new FileReader();
            reader.onloadend = () => {
              resolve(reader.result);
            };
            reader.readAsArrayBuffer(blob);
          }, "image/png");
        };
        img.src = e.target.result;
      };
      reader.readAsDataURL(file);
    });
  };

  const generateExcel = async () => {
    if (selectedFiles.length === 0) return;

    setIsProcessing(true);

    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet("Images");

      worksheet.columns = [
        { header: "Image Name", key: "name", width: 30 },
        { header: "Preview", key: "preview", width: 40 },
      ];

      worksheet.getRow(1).font = { bold: true };

      for (let i = 0; i < selectedFiles.length; i++) {
        const file = selectedFiles[i];
        const imageData = await createThumbnail(file);

        const rowIndex = i + 2;
        const row = worksheet.getRow(rowIndex);

        row.getCell(1).value = file.name;

        const imageId = workbook.addImage({
          buffer: imageData,
          extension: "png",
        });

        worksheet.addImage(imageId, {
          tl: { col: 1, row: rowIndex - 1 },
          ext: { width: 50, height: 50 },
        });

        row.height = 40;
      }

      const buffer = await workbook.xlsx.writeBuffer();

      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "images_report.xlsx";
      a.click();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      console.error("Error generating Excel:", error);
      alert("Error generating Excel file");
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <GradientBackground>
      <StyledPaper>
        <Typography
          variant="h4"
          gutterBottom
          align="center"
          sx={{
            fontFamily: "Montserrat, sans-serif",
            fontWeight: 600,
            marginBottom: "2rem",
          }}
        >
          Image Catalog to Excel
        </Typography>

        <BannerImage
          src="https://file.kikita.io/ipfs/bafkreiasxstg6yqwakolkxnxectwmv6spq6rp7h5yl7fdgt4gfop36zggi"
          alt="Banner"
        />

        <FileInput>
          <Typography
            variant="subtitle1"
            gutterBottom
            sx={{
              fontFamily: "Montserrat, sans-serif",
              marginBottom: "1rem",
            }}
          >
            This tool will help you organize all selected images into a single
            Excel file to create an image catalog document. The document format
            is .xls. Please select one or more images to archive in the catalog.
          </Typography>

          <input
            accept="image/*"
            style={{ display: "none" }}
            id="image-upload"
            multiple
            type="file"
            onChange={handleFileSelect}
          />

          <div
            style={{ display: "flex", justifyContent: "center", gap: "16px" }}
          >
            <label htmlFor="image-upload" style={{ width: "100%" }}>
              <UploadButton
                variant="contained"
                component="span"
                fullWidth
                startIcon={<CloudUploadIcon />}
              >
                Upload Images
              </UploadButton>
            </label>

            <GenerateButton
              variant="outlined"
              onClick={generateExcel}
              disabled={selectedFiles.length === 0 || isProcessing}
              startIcon={<FileDownloadIcon />}
              fullWidth
            >
              Generate Excel
            </GenerateButton>
          </div>
        </FileInput>

        {isProcessing && (
          <div
            style={{
              display: "flex",
              justifyContent: "center",
              margin: "20px 0",
            }}
          >
            <CircularProgress />
          </div>
        )}

        {selectedFiles.length > 0 && (
          <>
            <Typography
              variant="h6"
              gutterBottom
              sx={{
                fontFamily: "Montserrat, sans-serif",
                fontWeight: 500,
              }}
            >
              Selected Images ({selectedFiles.length}):
            </Typography>
            <List>
              {selectedFiles.map((file, index) => (
                <ListItem key={index}>
                  <ListItemText
                    primary={file.name}
                    secondary={`Size: ${(file.size / 1024).toFixed(2)} KB`}
                    sx={{
                      "& .MuiTypography-root": {
                        fontFamily: "Montserrat, sans-serif",
                      },
                    }}
                  />
                </ListItem>
              ))}
            </List>
          </>
        )}
      </StyledPaper>
    </GradientBackground>
  );
}

export default App;
