import { useEffect, useState } from "react";
import {
  Box,
  Button,
  Card,
  Flex,
  IconButton,
  Select,
  Text,
  Table,
  Thead,
  Tbody,
  Tr,
  Th,
  Td,
  Spinner,
  useToast,
  Modal,
  ModalOverlay,
  ModalContent,
  ModalHeader,
  ModalBody,
  ModalFooter,
  Input,
  Alert,
  AlertIcon,
} from "@chakra-ui/react";
import { FiFilter } from "react-icons/fi";
import FileManager from 'components/navbar/FileManager';


export default function CheckTable() {
  const [data, setData] = useState([]);
  const [filteredData, setFilteredData] = useState([]);
  const [columns, setColumns] = useState([]);
  const [loading, setLoading] = useState(false);
  const [selectedYear, setSelectedYear] = useState(null);
  const [selectedMonth, setSelectedMonth] = useState(null);
  const [filterOptions, setFilterOptions] = useState({});
  const [filterValues, setFilterValues] = useState({});
  const [showFilterPopup, setShowFilterPopup] = useState(false);
  const [currentFilterColumn, setCurrentFilterColumn] = useState(null);
  const [isFiltering, setIsFiltering] = useState(false);  // Loading state for filtering
  const toast = useToast();

  const generateYearRange = (startYear, endYear) => {
    const years = [];
    for (let year = startYear; year <= endYear; year++) {
      years.push(year);
    }
    return years;
  };

  const generateMonthRange = () => {
    return [
      "January", "February", "March", "April", "May", "June", 
      "July", "August", "September", "October", "November", "December"
    ];
  };

  const yearOptions = generateYearRange(2000, new Date().getFullYear());
  const monthOptions = generateMonthRange();

  const fetchTableData = async () => {
    if (!selectedYear || !selectedMonth) {
      toast({
        title: "Please select both year and month.",
        status: "info",
        duration: 3000,
        isClosable: true,
      });
      return;
    }

    setLoading(true);
    try {
      const response = await fetch(
        `http://localhost:8000/fetch-table-data/${selectedYear}/${selectedMonth}`
      );
      const result = await response.json();

      if (result.tableData && result.tableData.length > 0) {
        setData(result.tableData);
        setFilteredData(result.tableData);

        const dynamicColumns = Object.keys(result.tableData[0]).map((key) => ({
          accessorKey: key,
          header: key.toUpperCase(),
        }));
        setColumns(dynamicColumns);

        const options = {};
        dynamicColumns.forEach((col) => {
          options[col.accessorKey] = Array.from(
            new Set(result.tableData.map((item) => item[col.accessorKey]))
          );
        });
        setFilterOptions(options);

        if (result.tableData.length > 200) {
          toast({
            title: "Warning",
            description: "The data is too large!",
            status: "warning",
            duration: 5000,
            isClosable: true,
            position: "top",
          });
        }
      } else {
        toast({
          title: "No Data Found",
          status: "warning",
          duration: 3000,
          isClosable: true,
        });
      }
    } catch (error) {
      console.error("Error fetching table data:", error);
      toast({
        title: "Error Fetching Data",
        description: error.message,
        status: "error",
        duration: 5000,
        isClosable: true,
      });
    } finally {
      setLoading(false);
    }
  };

  const applyFilter = async () => {
    setIsFiltering(true);  // Set filtering loading state

    let filtered = data;

    Object.keys(filterValues).forEach((column) => {
      if (filterValues[column]) {
        filtered = filtered.filter((row) => {
          const value = row[column];
          if (value === undefined || value === null) {
            return false;  // Exclude row if value is undefined or null
          }
          return value.toString().toLowerCase().includes(filterValues[column].toLowerCase());
        });
      }
    });

    setFilteredData(filtered);
    setIsFiltering(false);  // Reset filtering loading state
  };

  const resetFilter = () => {
    setFilterValues({});
    setFilteredData(data);  // Reset to original data
  };

// Calculate the sum of the `quantity in su` field in filtered data
const calculateFilteredSum = () => {
    return filteredData.reduce((sum, row) => {
      return sum + (row["quantity in su"] || 0);  // Add quantity in su, default to 0 if not available
    }, 0);
  };
  
  return (
     <Box pt={{ base: "130px", md: "80px", xl: "80px" }}>
    <Card flexDirection="column" w="100%" px="0px" overflowX="auto" padding="12px">
      <Alert status="warning" mb="4">
        <AlertIcon />
        If the data is too large, May cause lag. Work slow for efficiency
      </Alert>

      <Flex px="25px" mb="8px" justifyContent="space-between" align="center">
        <Text fontSize="22px" fontWeight="700">
          Data Table
        </Text>

        <Flex align="center" gap="4">
        <FileManager />
          <Select
            placeholder="Select Year"
            width="150px"
            onChange={(e) => setSelectedYear(e.target.value)}
            value={selectedYear || ""}
          >
            {yearOptions.map((year) => (
              <option key={year} value={year}>
                {year}
              </option>
            ))}
          </Select>
          <Select
            placeholder="Select Month"
            width="150px"
            onChange={(e) => setSelectedMonth(e.target.value)}
            value={selectedMonth || ""}
          >
            {monthOptions.map((month, index) => (
              <option key={index} value={month}>
                {month}
              </option>
            ))}
          </Select>
          <Button  colorScheme="teal" onClick={fetchTableData} borderRadius="5">
            Fetch Data
          </Button>
        </Flex>
      </Flex>

      {/* Display Totals */}
      <Box mb="4">
        <Text fontSize="16px" fontWeight="600" mb="2">
          Total Records: {data.length}
        </Text>
        <Text fontSize="16px" fontWeight="600" mb="2">
          Filtered Records: {filteredData.length}
        </Text>
        <Text fontSize="16px" fontWeight="600" mb="2">
          Filtered Sum of Quantity: {calculateFilteredSum()}
        </Text>
      </Box>

      {/* Filters Section */}
      <Box mb="4" >
        <Text fontSize="16px" fontWeight="600" mb="2" >
          Apply Filters
        </Text>
        <Flex mb="4" gap="4" >
          {columns.map((column) => (
            <Select
              key={column.accessorKey}
              placeholder={`Filter by ${column.header}`}
              value={filterValues[column.accessorKey] || ""}
              onChange={(e) => {
                setFilterValues({
                  ...filterValues,
                  [column.accessorKey]: e.target.value,
                });
              }}
            >
              {filterOptions[column.accessorKey]?.map((option) => (
                <option key={option} value={option}>
                  {option}
                </option>
              ))}
            </Select>
          ))}
        </Flex>
        <Button
          colorScheme="teal"
          onClick={applyFilter}
          isLoading={isFiltering}
          loadingText="Applying Filter"
        >
          Apply Filter
        </Button>
        <Button ml="4" onClick={resetFilter}>
          Reset Filter
        </Button>
      </Box>

      <Box>
        {loading ? (
          <Flex justify="center" align="center" height="300px">
            <Spinner size="xl" />
          </Flex>
        ) : (
          <Table variant="simple" color="gray.500" mb="24px" mt="12px" >
            <Thead >
              <Tr>
                {columns.map((column) => (
                  <Th key={column.accessorKey} padding="12px">
                    <Flex align="center">
                      {column.header}
                      <IconButton
                        ml="2"
                        size="sm"
                        icon={<FiFilter />}
                        onClick={() => {
                          setCurrentFilterColumn(column.accessorKey);
                          setShowFilterPopup(true);
                        }}
                        aria-label={`Filter ${column.header}`}
                      />
                    </Flex>
                  </Th>
                ))}
              </Tr>
            </Thead>
            <Tbody>
              {filteredData.length > 0 ? (
                filteredData.map((row, rowIndex) => (
                  <Tr key={rowIndex}>
                    {columns.map((column) => (
                      <Td key={column.accessorKey} padding="12px">{row[column.accessorKey]} </Td>
                    ))}
                  </Tr>
                ))
              ) : (
                <Tr>
                  <Td colSpan={columns.length}>No data available</Td>
                </Tr>
              )}
            </Tbody>
          </Table>
        )}
      </Box>

      <Modal isOpen={showFilterPopup} onClose={resetFilter}>
        <ModalOverlay />
        <ModalContent>
          <ModalHeader>Filter by {currentFilterColumn}</ModalHeader>
          <ModalBody>
            <Input
              placeholder={`Enter ${currentFilterColumn} value to filter`}
              onChange={(e) => {
                setFilterValues({
                  ...filterValues,
                  [currentFilterColumn]: e.target.value,
                });
              }}
              value={filterValues[currentFilterColumn] || ""}
            />
          </ModalBody>

          <ModalFooter>
            <Button variant="ghost" onClick={resetFilter}>
              Close
            </Button>
            <Button colorScheme="teal" onClick={applyFilter}>
              Apply Filter
            </Button>
          </ModalFooter>
        </ModalContent>
      </Modal>
    </Card>
    </Box>
  );
}
