import React, { useEffect, useState } from "react";
import {
  Box,
  Radio,
  RadioGroup,
  Stack,
  Text,
  useColorModeValue,
} from "@chakra-ui/react";
import { Bar } from "react-chartjs-2";
import axios from "axios";
import { Chart as ChartJS, CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend } from "chart.js"; // Import necessary components
import LoadingModal from "./components/LoadingModal";

// Register necessary components for chart.js
ChartJS.register(CategoryScale, LinearScale, BarElement, Title, Tooltip, Legend);

export default function FetchDataPage() {
  const [data, setData] = useState({});
  const [selectedOption, setSelectedOption] = useState("sales-group");
  const [dataType, setDataType] = useState("lubricants"); // 'lubricants' or 'petroleum'
  const [isLoading, setIsLoading] = useState(false);

  const brandColor = useColorModeValue("brand.500", "white");
  const boxBg = useColorModeValue("white", "gray.700");
  const axisColor = useColorModeValue("black", "white");

  const fetchData = async () => {
    setIsLoading(true);
    let url = "";
    try {
      if (selectedOption === "sales-group") {
        url = `http://localhost:8000/fetch-data-by-sales-group?type=${dataType}`;
      } else if (selectedOption === "customer-code") {
        url = `http://localhost:8000/fetch-data-by-customer-code?type=${dataType}`;
      } else if (selectedOption === "material-code") {
        url = `http://localhost:8000/fetch-data-by-material-code?type=${dataType}`;
      }

      const response = await axios.get(url);
      setData(response.data);
    } catch (error) {
      console.error("Error fetching data:", error.message);
    } finally {
      setIsLoading(false);
    }
  };

  useEffect(() => {
    fetchData();
  }, [selectedOption, dataType]);

  const labels = Object.keys(data);
  const values = Object.values(data);

  // Dynamically calculate the chart's width
  const chartWidth = Math.max(labels.length * 50, 800); // Minimum width: 800px, ~50px per bar

  const chartData = {
    labels,
    datasets: [
      {
        label: `Data for ${dataType} - ${selectedOption.replace("-", " ")}`,
        data: values,
        backgroundColor: "rgba(75,192,192,0.6)",
        borderColor: "rgba(255,255,255,1)",
        hoverBackgroundColor: "rgba(0,0,0,0.2)",
        borderWidth: 2,
        borderRadius: 10,
        barPercentage: 0.9, // Wider bars
        categoryPercentage: 0.4, // Reduced spacing
      },
    ],
  };

  const chartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        labels: {
          color: axisColor,
          font: {
            size: 14,
            family: "Poppins",
          },
        },
      },
      tooltip: {
        callbacks: {
          title: function (tooltipItems) {
            // Set the title to the material code
            return `Material Code: ${tooltipItems[0].label}`;
          },
          label: function (tooltipItem) {
            // Set the label to display the sales value
            return `Sales: ${tooltipItem.raw.toLocaleString()}`;
          },
        },
        backgroundColor: "rgba(0, 0, 0, 0.8)",
        titleFont: { size: 14, weight: "bold" },
        bodyFont: { size: 12 },
        cornerRadius: 5,
        padding: 10,
      },
    },
    scales: {
      x: {
        ticks: {
          color: axisColor,
          font: { size: 12, family: "Poppins" },
        },
        grid: { display: false },
      },
      y: {
        beginAtZero: true,
        ticks: {
          
          callback: function (value) {
            return value.toLocaleString();
          },
          color: axisColor,
          font: { size: 12, family: "Poppins" },
        },
        grid: { color: "rgba(200,200,200,0.2)" },
      },
    },
    datasets: {
      bar: {
        minBarLength: 5,
      },
    },
  };
  
  return (
    <Box pt={{ base: "130px", md: "80px", xl: "80px" }}>
      <LoadingModal isOpen={isLoading} />

      {/* Data Type Selector */}
      <Box mb="20px" p="20px" borderRadius="15px" bg={boxBg} boxShadow="md">
        <Text fontSize="lg" fontWeight="bold" mb="10px" color={brandColor}>
          Select Data Type:
        </Text>
        <RadioGroup
          onChange={(value) => setDataType(value)}
          value={dataType}
          colorScheme="blue"
        >
          <Stack direction="row" spacing={6}>
            <Radio value="lubricants">Lubricants</Radio>
            <Radio value="petroleum">Petroleum</Radio>
          </Stack>
        </RadioGroup>
      </Box>

      {/* Data Fetch Selector */}
      <Box mb="20px" p="20px" borderRadius="15px" bg={boxBg} boxShadow="md">
        <Text fontSize="lg" fontWeight="bold" mb="10px" color={brandColor}>
          Select Data to Fetch:
        </Text>
        <RadioGroup
          onChange={(value) => setSelectedOption(value)}
          value={selectedOption}
          colorScheme="blue"
        >
          <Stack direction="row" spacing={6}>
            <Radio value="sales-group">Sales Group</Radio>
            <Radio value="customer-code">Customer Code</Radio>
            <Radio value="material-code">Material Code</Radio>
          </Stack>
        </RadioGroup>
      </Box>

      {/* Scrollable Chart */}
      <Box
        p="20px"
        borderRadius="15px"
        bg={boxBg}
        boxShadow="md"
        overflowX="auto"
        whiteSpace="nowrap"
      >
        <Text fontSize="lg" fontWeight="bold" mb="10px" color={brandColor}>
          Aggregated Data: {dataType.toUpperCase()} - {selectedOption.replace("-", " ").toUpperCase()}
        </Text>
        <Box height="500px" minWidth={`${chartWidth}px`} overflow="hidden">
          <Bar data={chartData} options={chartOptions} />
        </Box>
      </Box>
    </Box>
  );
}
