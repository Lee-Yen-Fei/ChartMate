import React, { useState, useRef } from "react";
import { Button, Rows, Text, Select, Box, NumberInput } from "@canva/app-ui-kit";
import { addNativeElement, getCurrentPageContext } from "@canva/design";
import { upload } from "@canva/asset";
import styles from "styles/components.css";
import ExcelJS from "exceljs";
import { Chart, ChartTypeRegistry, ChartConfiguration } from "chart.js/auto";

export const App = () => {
  const [file, setFile] = useState<File | null>(null);
  const [fileName, setFileName] = useState("");
  const [chartType, setChartType] = useState<keyof ChartTypeRegistry>("bar");
  const [dataRange, setDataRange] = useState<string>("A1:B10");
  const [chartData, setChartData] = useState<any[]>([]);
  const chartRef = useRef<Chart | null>(null);
  const [fileUrl, setFileUrl] = useState<string>("");
  const [summaryType, setSummaryType] = useState<'text' | 'statistics'>('text');

  const handleUrlChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    setFileUrl(e.target.value);
  };

  const loadFileFromUrl = async (url: string) => {
    try {
      const response = await fetch(url);
      const arrayBuffer = await response.arrayBuffer();
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(arrayBuffer);
      return workbook;
    } catch (error) {
      console.error("Failed to load .xlsx file from URL:", error);
      return null;
    }
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
      setFileName(e.target.files[0].name);
    }
  };

  const handleGenerateChart = async () => {
    let workbook: ExcelJS.Workbook | null = null;
  
    if (file) {
      workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());
    } else if (fileUrl) {
      workbook = await loadFileFromUrl(fileUrl);
    }
  
    if (!workbook) {
      console.error("No workbook found");
      return;
    }
  
    const worksheet = workbook.getWorksheet(1);
  
    if (!worksheet) {
      console.error("Worksheet not found");
      return;
    }
  
    const data: any[] = [];
  
    // Extract headers from the first row
    const headers = worksheet.getRow(1).values as any[];
  
    // Determine the columns for x and y axes based on the data range
    const [startCell, endCell] = dataRange.split(":");
    const startColumn = startCell.replace(/\d/g, ''); // Extract column letter (e.g., "B")
    const endColumn = endCell.replace(/\d/g, ''); // Extract column letter (e.g., "C")
    const startRow = parseInt(startCell.replace(/\D/g, ""), 10); // Extract row number (e.g., 1)
    const endRow = parseInt(endCell.replace(/\D/g, ""), 10); // Extract row number (e.g., 5)
  
    // Find columns for x and y data
    const xAxisColumnIndex = worksheet.getCell(startColumn + '1').col;
    const yAxisColumnIndex = worksheet.getCell(endColumn + '1').col;
  
    for (let i = startRow; i <= endRow; i++) {
      const row = worksheet.getRow(i);
      const xValue = row.getCell(xAxisColumnIndex).value;
      const yValue = row.getCell(yAxisColumnIndex).value;
  
      if (xValue && yValue) {
        data.push([xValue, yValue]);
      }
    }
  
    console.log("Extracted Data:", data); // Debugging: Check extracted data
  
    setChartData(data);
    renderChart(data, headers);
  };  

  const renderChart = (data: any[], headers: any[]) => {
    const ctx = document.getElementById("chart") as HTMLCanvasElement;
  
    if (chartRef.current) {
      chartRef.current.destroy();
    }
  
    if (ctx) {
      const generateRandomColor = () =>
        `rgba(${Math.floor(Math.random() * 255)}, ${Math.floor(Math.random() * 255)}, ${Math.floor(Math.random() * 255)}, 0.8)`; // Adjusted opacity
      const generateRandomBorderColor = () =>
        `rgba(${Math.floor(Math.random() * 255)}, ${Math.floor(Math.random() * 255)}, ${Math.floor(Math.random() * 255)}, 1)`;
  
      const labels = data.map((row) => row[0]);
      const datasetData = data.map((row) => row[1]);
  
      const chartConfig: ChartConfiguration<keyof ChartTypeRegistry> = {
        type: chartType,
        data: {
          labels: labels,
          datasets: [
            {
              label: headers[2] || "Dataset", // Use the second header as dataset label
              data: datasetData,
              backgroundColor:
                chartType === "bar" || chartType === "pie" || chartType === "doughnut"
                  ? datasetData.map(() => generateRandomColor())
                  : undefined,
              borderColor:
                chartType === "bar"
                  ? datasetData.map(() => generateRandomBorderColor())
                  : undefined,
              borderWidth: chartType === "bar" ? 1 : undefined,
              borderRadius: chartType === "bar" ? 5 : undefined,
            },
          ],
        },
        options: {
          responsive: true,
          plugins: {
            legend: {
              position: "top" as const,
              labels: {
                font: {
                  size: 10, // Adjust legend font size
                },
              },
            },
            tooltip: {
              callbacks: {
                label: function (context) {
                  return `${context.label}: ${context.raw}`;
                },
              },
            },
          },
          scales: {
            x: {
              type: 'category',
              title: {
                display: true,
                text: headers[1] || 'X-Axis Label', // Use the first header as x-axis title
                font: {
                  size: 12, // Adjust x-axis title font size
                },
              },
              ticks: {
                font: {
                  size: 10, // Adjust x-axis label font size
                },
              },
            },
            y: {
              title: {
                display: true,
                text: headers[2] || 'Y-Axis Label', // Use the second header as y-axis title
                font: {
                  size: 12, // Adjust y-axis title font size
                },
              },
              ticks: {
                font: {
                  size: 10, // Adjust y-axis label font size
                },
              },
            },
          },
        },
      };
  
      chartRef.current = new Chart(ctx, chartConfig);
    }
  };  

  const exportChartToCanva = async () => {
  try {
    const chart = chartRef.current;
    if (!chart) {
      console.error("No chart instance found");
      return;
    }

    const imageData = chart.toBase64Image();
    const result = await upload({
      type: "IMAGE",
      mimeType: "image/png",
      url: imageData,
      thumbnailUrl: imageData,
    });

    await addNativeElement({
      type: "IMAGE",
      ref: result.ref,
    });

    console.log("Chart image added to Canva design successfully.");
  } catch (error) {
    console.error("Error exporting chart to Canva:", error);
  }
};

  const generateSummary = async () => {
    if (chartData.length === 0) {
      alert("No data available to summarize.");
      return;
    }
    
    const context = await getCurrentPageContext();

    if (!context.dimensions) {
      console.warn("The current design does not have dimensions");
      return;
    }

    console.log("Generating summary for data:", chartData);

    const labels = chartData.map((row) => String(row[0]));
    const values = chartData.map((row) => {
      const value = parseFloat(row[1]);
      return isNaN(value) ? 0 : value;
    });

    const total = values.reduce((acc, val) => acc + val, 0);
    const maxValue = Math.max(...values);
    const maxLabel = labels[values.indexOf(maxValue)];
    const minValue = Math.min(...values);
    const minLabel = labels[values.indexOf(minValue)];
    const mean = (Math.round(total / values.length * 100) / 100).toFixed(2);

    const sortedValues = [...values].sort((a, b) => a - b);
    const middleIndex = Math.floor(sortedValues.length / 2);
    const median = sortedValues.length % 2 === 0
      ? (sortedValues[middleIndex - 1] + sortedValues[middleIndex]) / 2
      : sortedValues[middleIndex];

    const frequencyMap: { [key: number]: number } = {};
    values.forEach((value) => {
      frequencyMap[value] = (frequencyMap[value] || 0) + 1;
    });
    const maxFrequency = Math.max(...Object.values(frequencyMap));
    const modes = Object.keys(frequencyMap)
      .filter((key) => frequencyMap[parseFloat(key)] === maxFrequency)
      .map(Number);

    let summary = "";

    if (summaryType === 'text') {
      summary = `This ${chartType} chart represents data with a total sum of ${total}. The mean is ${mean}, median is ${median}, and mode is ${modes}. The highest value is ${maxValue}, represented by ${maxLabel}. The smallest value is ${minValue}, represented by ${minLabel}.`;
    } else if (summaryType === 'statistics') {
      summary = `Total sum: ${total}\nMean: ${mean}\nMedian: ${median}\nMode: ${modes.join(", ")} (appears ${maxFrequency} times)\nHighest value: ${maxValue} (${maxLabel})\nSmallest value: ${minValue} (${minLabel})`;
    }

    console.log("Generated Summary:", summary);

    alert(`Generated Summary: ${summary}`);

    const width = 500;
    const height = 100;
    const top = context.dimensions.height - height;
    const left = context.dimensions.width - width;

    await addNativeElement({
      type: "TEXT",
      width,
      top,
      left,
      rotation: 0,
      children: [summary],
    });
  };

  return (
    <div className={styles.scrollContainer}>
      <Rows spacing="2u">
        <Text>
          Upload an Excel file, choose your chart type, and generate the chart.
        </Text>
        <input type="file" onChange={handleFileChange} accept=".xlsx" />
        <Text>or</Text>
        <input
          type="text"
          placeholder="Enter URL of .xlsx file"
          value={fileUrl}
          onChange={handleUrlChange}
        />
        <Box>
          <Select
            value={chartType}
            onChange={(value) => setChartType(value as keyof ChartTypeRegistry)}
            options={[
              { value: "bar", label: "Bar" },
              { value: "line", label: "Line" },
              { value: "pie", label: "Pie" },
              { value: "doughnut", label: "Doughnut" },
            ]}
          />
        </Box>

        <Box>
          <Text>Data Range (e.g., A1:B10)</Text>
          <input
            type="text"
            value={dataRange}
            onChange={(e) => setDataRange(e.target.value)}
            placeholder="Data Range (e.g., A1:B10)"
          />
        </Box>
        <Button
          variant="primary"
          onClick={handleGenerateChart}
          stretch
        >
          Generate Chart
        </Button>

        <canvas id="chart" style={{ maxWidth: "100%", marginTop: "20px" }}></canvas>

        <Button
          variant="secondary"
          onClick={exportChartToCanva}
          stretch
        >
          Export Chart to Canva
        </Button>

        <Box>
          <Text>Summary Type</Text>
          <Select
            value={summaryType}
            onChange={(value) => setSummaryType(value as 'text' | 'statistics')}
            options={[
              { value: 'text', label: 'Text Summary' },
              { value: 'statistics', label: 'Summarized Statistics' },
            ]}
          />
        </Box>

        <Button
          variant="primary"
          onClick={generateSummary}
          stretch
        >
          Generate Summary
        </Button>
      </Rows>
    </div>
  );
};
