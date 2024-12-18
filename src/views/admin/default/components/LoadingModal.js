import React, { useEffect, useState } from "react";
import {
  Modal,
  ModalOverlay,
  ModalContent,
  ModalBody,
  ModalFooter,
  Center,
  Spinner,
  Text,
} from "@chakra-ui/react";

const LoadingModal = ({ isOpen }) => {
  const [loadingText, setLoadingText] = useState("Please wait, calculating data...");

  useEffect(() => {
    const timer = setTimeout(() => {
      setLoadingText("Please wait! Data is too long...");
    }, 10000); // Change text after 10 seconds

    return () => clearTimeout(timer); // Cleanup the timer
  }, []);

  return (
    <Modal isOpen={isOpen} onClose={() => {}} isCentered>
      <ModalOverlay />
      <ModalContent>
        <ModalBody>
          <Center py={8} flexDirection="column">
            <Spinner size="xl" />
            <Text mt={4} fontSize="lg" fontWeight="bold">
              {loadingText}
            </Text>
          </Center>
        </ModalBody>
        <ModalFooter />
      </ModalContent>
    </Modal>
  );
};

export default LoadingModal;
