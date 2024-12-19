import React, { useState, useEffect } from "react";
import {
  Box,
  Button,
  Modal,
  ModalOverlay,
  ModalContent,
  ModalHeader,
  ModalCloseButton,
  ModalBody,
  Text,
  VStack,
  HStack,
  useColorModeValue,
  Spinner,
} from "@chakra-ui/react";
import axios from "axios";

const FileManager = () => {
  const [showModal, setShowModal] = useState(false);
  const [files, setFiles] = useState({ lubricants: [], petroleum: [] });
  const [loading, setLoading] = useState(false);
  const [currentOperation, setCurrentOperation] = useState(null); // Tracks specific operations

  const bgColor = useColorModeValue("white", "gray.800");
  const textColor = useColorModeValue("black", "white");
  const headerBg = useColorModeValue("gray.100", "gray.700");
  const btnColor = useColorModeValue("blue.500", "blue.200");
  const btnTextColor = useColorModeValue("white", "black");

  // Fetch files from the server
  const fetchFiles = async () => {
    try {
      setLoading(true);
      const response = await axios.get(`http://localhost:8000/fetch-files`);
      setFiles(response.data);
    } catch (error) {
      console.error("Error fetching files:", error);
    } finally {
      setLoading(false);
    }
  };

  // Delete a specific file
  const deleteFile = async (fileName, type) => {
    setCurrentOperation(`Deleting ${fileName} from ${type}...`);
    try {
      await axios.delete(
        `http://localhost:8000/delete-file?fileType=${type}&fileName=${fileName}`
      );
      fetchFiles();
    } catch (error) {
      console.error("Error deleting file:", error);
    } finally {
      setCurrentOperation(null);
    }
  };

  // Delete all files in a specific category
  const deleteAllFiles = async (type) => {
    setCurrentOperation(`Deleting all files from ${type}...`);
    try {
      await axios.delete(`http://localhost:8000/delete-all-files?fileType=${type}`);
      fetchFiles();
    } catch (error) {
      console.error("Error deleting all files:", error);
    } finally {
      setCurrentOperation(null);
    }
  };

  useEffect(() => {
    if (showModal) fetchFiles();
  }, [showModal]);

  return (
    <>
      <Button
        bg={btnColor}
        color={btnTextColor}
        onClick={() => setShowModal(true)}
      >
        Manage Files
      </Button>

      <Modal isOpen={showModal} onClose={() => setShowModal(false)} size="xl">
        <ModalOverlay />
        <ModalContent bg={bgColor} color={textColor}>
          <ModalHeader bg={headerBg} borderRadius="10px 10px 0 0">
            File Manager
          </ModalHeader>
          <ModalCloseButton />
          <ModalBody>
            {loading ? (
              <VStack>
                <Spinner size="lg" />
                <Text>Loading files...</Text>
              </VStack>
            ) : (
              <>
                <FileSection
                  title="Lubricants"
                  files={files.lubricants}
                  onDelete={(fileName) => deleteFile(fileName, "lubricants")}
                  onDeleteAll={() => deleteAllFiles("lubricants")}
                  isProcessing={currentOperation?.includes("lubricants")}
                />

                <FileSection
                  title="Petroleum"
                  files={files.petroleum}
                  onDelete={(fileName) => deleteFile(fileName, "petroleum")}
                  onDeleteAll={() => deleteAllFiles("petroleum")}
                  isProcessing={currentOperation?.includes("petroleum")}
                />
              </>
            )}
            {currentOperation && (
              <Text mt={4} fontStyle="italic" textAlign="center">
                {currentOperation}
              </Text>
            )}
          </ModalBody>
        </ModalContent>
      </Modal>
    </>
  );
};

// Sub-component for displaying a file section
const FileSection = ({ title, files, onDelete, onDeleteAll, isProcessing }) => {
  const sectionBg = useColorModeValue("gray.100", "gray.700");

  return (
    <Box mb={4} p={4} bg={sectionBg} borderRadius="md">
      <HStack justify="space-between" mb={4}>
        <Text fontWeight="bold">{title}</Text>
        <Button
          size="sm"
          colorScheme="red"
          onClick={onDeleteAll}
          isDisabled={isProcessing}
        >
          {isProcessing ? "Processing..." : "Delete All"}
        </Button>
      </HStack>
      <VStack align="stretch" spacing={2}>
        {files.length > 0 ? (
          files.map((file) => (
            <HStack
              key={file.name}
              justify="space-between"
              bg="gray.50"
              _dark={{ bg: "gray.600" }}
              p={2}
              borderRadius="md"
            >
              <Text>{file.name}</Text>
              <Button
                size="sm"
                colorScheme="red"
                onClick={() => onDelete(file.name)}
                isDisabled={isProcessing}
              >
                {isProcessing ? "Processing..." : "Delete"}
              </Button>
            </HStack>
          ))
        ) : (
          <Text>No files available in this section.</Text>
        )}
      </VStack>
    </Box>
  );
};

export default FileManager;
