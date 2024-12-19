import React, { useEffect, useState } from "react";
import {
  Box,
  SimpleGrid,
 
  Icon,
  Text,
  useColorModeValue,
} from "@chakra-ui/react";
import MiniStatistics from "components/card/MiniStatistics";
import IconBox from "components/icons/IconBox";
import { MdPeople, MdShoppingCart, MdAccountBalanceWallet } from "react-icons/md";
import { Bar, Line } from "react-chartjs-2";
import axios from "axios";
import LoadingModal from "./components/LoadingModal";
import CheckTable from './components/CheckTable'; // Adjust the path based on your file structure


// Register Chart.js components
import {
  Chart as ChartJS,
  CategoryScale,
  LinearScale,
  BarElement,
  LineElement,
  PointElement,
  Title,
  Tooltip,
  Legend,
} from "chart.js";

ChartJS.register(CategoryScale, LinearScale, BarElement, LineElement, PointElement, Title, Tooltip, Legend);

// Mini Statistics Component
const Statistics = ({ data, brandColor, boxBg, textColor }) => (
  <SimpleGrid columns={{ base: 1, md: 2, lg: 3 }} gap="20px" mb="20px">
    <MiniStatistics
      startContent={
        <IconBox
          w="56px"
          h="56px"
          bg={boxBg}
          icon={<Icon w="32px" h="32px" as={MdPeople} color={brandColor} />}
        />
      }
      name="Total Users"
      value={data.totalUsers}
      textColor={textColor}
    />
    <MiniStatistics
      startContent={
        <IconBox
          w="56px"
          h="56px"
          bg={boxBg}
          icon={<Icon w="32px" h="32px" as={MdShoppingCart} color={brandColor} />}
        />
      }
      name="Total Orders"
      value={data.totalOrders}
      textColor={textColor}
    />
    <MiniStatistics
      startContent={
        <IconBox
          w="56px"
          h="56px"
          bg={boxBg}
          icon={<Icon w="32px" h="32px" as={MdAccountBalanceWallet} color={brandColor} />}
        />
      }
      name="Total Sales"
      value={`$${data.totalSales.toFixed(2)}`}
      textColor={textColor}
    />
  </SimpleGrid>
);

// Chart Components
const ChartCard = ({ title, children, boxBg, brandColor }) => (
  <Box p="20px" borderRadius="15px" bg={boxBg} boxShadow="md">
    <Text fontSize="lg" fontWeight="bold" mb="10px" color={brandColor}>
      {title}
    </Text>
    {children}
  </Box>
);

const YearlySalesBarChart = ({ yearlySales, brandColor, axisColor }) => {
  const chartData = {
    labels: Object.keys(yearlySales),
    datasets: [
      {
        label: "Yearly Sales",
        data: Object.values(yearlySales),
        backgroundColor: brandColor,
        borderWidth: 1,
      },
    ],
  };

  const chartOptions = {
    responsive: true,
    scales: {
      x: {
        ticks: {
          color: axisColor, // Set axis tick color for visibility
        },
      },
      y: {
        ticks: {
          color: axisColor, // Set axis tick color for visibility
        },
      },
    },
  };

  return <Bar data={chartData} options={chartOptions} />;
};

const YearlySalesLineChart = ({ yearlySales, brandColor, axisColor }) => {
  const chartData = {
    labels: Object.keys(yearlySales),
    datasets: [
      {
        label: "Yearly Sales",
        data: Object.values(yearlySales),
        borderColor: brandColor,
        backgroundColor: "transparent",
        tension: 0.4,
        pointRadius: 5,
      },
    ],
  };

  const chartOptions = {
    responsive: true,
    scales: {
      x: {
        ticks: {
          color: axisColor, // Set axis tick color for visibility
        },
      },
      y: {
        ticks: {
          color: axisColor, // Set axis tick color for visibility
        },
      },
    },
  };

  return <Line data={chartData} options={chartOptions} />;
};

// Main Component
export default function UserReports() {
  const [data, setData] = useState({
    totalUsers: 0,
    totalOrders: 0,
    totalSales: 0,
  });
  const [yearlySales, setYearlySales] = useState({});
  const [isLoading, setIsLoading] = useState(true);

  const brandColor = useColorModeValue("brand.500", "white");
  const boxBg = useColorModeValue("white", "gray.700");
  const textColor = useColorModeValue("black", "white");
  const axisColor = useColorModeValue("black", "white"); // Ensure axis ticks are visible in dark mode

  useEffect(() => {
    const fetchData = async () => {
      try {
        setIsLoading(true);
        const response = await axios.get("http://localhost:8000/fetch-excel-data");
        const yearlySalesResponse = await axios.get("http://localhost:8000/fetch-sales-by-year");

        setData(response.data);
        setYearlySales(yearlySalesResponse.data);
      } catch (error) {
        console.error("Error fetching data:", error);
      } finally {
        setIsLoading(false);
      }
    };

    fetchData();
  }, []);

  return (
    <Box pt={{ base: "130px", md: "80px", xl: "80px" }}>
      <LoadingModal isOpen={isLoading} />
      {!isLoading && (
        <>
          {/* Mini Statistics */}
          <Statistics data={data} brandColor={brandColor} boxBg={boxBg} textColor={textColor} />

          {/* Charts Section */}
          <SimpleGrid columns={{ base: 1, md: 2 }} gap="20px" mb="20px">
            <ChartCard title="Yearly Sales Bar Chart" boxBg={boxBg} brandColor={brandColor}>
              <YearlySalesBarChart yearlySales={yearlySales} brandColor={brandColor} axisColor={axisColor} />
            </ChartCard>
            <ChartCard title="Yearly Sales Line Chart" boxBg={boxBg} brandColor={brandColor}>
              <YearlySalesLineChart yearlySales={yearlySales} brandColor={brandColor} axisColor={axisColor} />
            </ChartCard>
          </SimpleGrid>
          
        </>
      )}
    </Box>
  );
}
