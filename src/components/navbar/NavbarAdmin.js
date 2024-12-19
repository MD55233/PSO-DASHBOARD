// Chakra Imports
import {
	Box,
	Breadcrumb,
	BreadcrumbItem,
	BreadcrumbLink,
	Flex,
	Text,
	Link,
	Button,
	Modal,
	ModalOverlay,
	ModalContent,
	ModalHeader,
	ModalCloseButton,
	ModalBody,
	ModalFooter,
	Input,
	useColorModeValue,
	useDisclosure,
	useToast,
	Container
} from '@chakra-ui/react';
import PropTypes from 'prop-types';
import React, { useState, useEffect } from 'react';
import AdminNavbarLinks from 'components/navbar/NavbarLinksAdmin';
import * as XLSX from 'xlsx';
import axios from 'axios';
import FileManager from './FileManager';

export default function AdminNavbar(props) {
	const [scrolled, setScrolled] = useState(false);
	const [selectedFiles, setSelectedFiles] = useState([]); // Selected Files
	const [validatedFiles, setValidatedFiles] = useState([]); // Files After Validation
	const [isValidating, setIsValidating] = useState(false); // Loading State for Validation
	const { isOpen, onOpen, onClose } = useDisclosure();
	const toast = useToast();

	useEffect(() => {
		window.addEventListener('scroll', changeNavbar);
		return () => window.removeEventListener('scroll', changeNavbar);
	}, []);


	const { brandText } = props;

	const handleFileChange = (e) => {
		const files = Array.from(e.target.files);
		setSelectedFiles(files);
		setValidatedFiles([]); // Reset validated files
	};

	// Function to normalize headers (trim spaces and lower case)
	const normalizeHeaders = (headers) => {
		return headers.map((header) => header.trim().toLowerCase());
	};

	const validateHeaders = async () => {
		const expectedHeaders1 = [
			'sales grp', 'customer code', 'customer name',
			'material code', 'shipping point name', 'vehicle text',
			'billing date', 'quantity in su'
		].map(h => h.trim().toLowerCase());

		const expectedHeaders2 = [
			'sales grp', 'customer code', 'customer name',
			'material name', 'material code', 'quantity in su',
			'billing date', 'sku qty'
		].map(h => h.trim().toLowerCase());

		setIsValidating(true);
		let validFiles = [];

		for (const file of selectedFiles) {
			let fileIsValid = false;

			try {
				const data = await file.arrayBuffer();
				const workbook = XLSX.read(data, { type: 'array' });

				for (const sheetName of workbook.SheetNames) {
					const worksheet = workbook.Sheets[sheetName];
					const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

					if (!sheetData || sheetData.length === 0) {
						continue;
					}

					const headers = sheetData[0].map((h) => h?.toString().trim().toLowerCase());
					const isValid = expectedHeaders1.every(h => headers.includes(h)) ||
									expectedHeaders2.every(h => headers.includes(h));

					if (isValid) {
						fileIsValid = true;
						break;
					}
				}

				if (fileIsValid) {
					validFiles.push(file);
				} else {
					toast({
						title: 'Invalid File',
						description: `${file.name} does not have valid headers.`,
						status: 'error',
						duration: 3000,
						isClosable: true
					});
				}
			} catch (error) {
				console.error('Error processing file:', error);
				toast({
					title: 'File Error',
					description: `Unable to process ${file.name}.`,
					status: 'error',
					duration: 3000,
					isClosable: true
				});
			}
		}

		setIsValidating(false);

		if (validFiles.length) {
			setValidatedFiles(validFiles);
			toast({
				title: 'Files Validated',
				description: `${validFiles.length} file(s) passed validation.`,
				status: 'success',
				duration: 3000,
				isClosable: true
			});
		} else {
			toast({
				title: 'Validation Failed',
				description: 'No files passed validation.',
				status: 'warning',
				duration: 3000,
				isClosable: true
			});
		}
	};

	const uploadFiles = async (type) => {
		if (!validatedFiles.length) {
			toast({
				title: 'No Files to Upload',
				description: 'Please validate files before uploading.',
				status: 'warning',
				duration: 3000,
				isClosable: true
			});
			return;
		}
	
		const formData = new FormData();
		validatedFiles.forEach((file) => formData.append('files', file));
	
		try {
			// Pass fileType as a query parameter
			await axios.post(`http://localhost:8000/upload-excel?fileType=${type}`, formData, {
				headers: {
					'Content-Type': 'multipart/form-data'
				}
			});
	
			toast({
				title: 'Upload Successful',
				description: `Files uploaded successfully to ${type}.`,
				status: 'success',
				duration: 3000,
				isClosable: true
			});
	
			setValidatedFiles([]); // Clear files after upload
			onClose();
		} catch (error) {
			console.error('Upload Error:', error);
			toast({
				title: 'Upload Failed',
				description: 'An error occurred while uploading files.',
				status: 'error',
				duration: 3000,
				isClosable: true
			});
		}
	};

	useEffect(() => {
		window.addEventListener('scroll', changeNavbar);

		return () => {
			window.removeEventListener('scroll', changeNavbar);
		};
	});

	const { secondary, message } = props;

	// Here are all the props that may change depending on navbar's type or state.(secondary, variant, scrolled)
	let mainText = useColorModeValue('navy.700', 'white');
	let secondaryText = useColorModeValue('gray.700', 'white');
	let navbarPosition = 'fixed';
	let navbarFilter = 'none';
	let navbarBackdrop = 'blur(20px)';
	let navbarShadow = 'none';
	let navbarBg = useColorModeValue('rgba(244, 247, 254, 0.2)', 'rgba(11,20,55,0.5)');
	let navbarBorder = 'transparent';
	let secondaryMargin = '0px';
	let paddingX = '15px';
	let gap = '0px';
	const changeNavbar = () => {
		if (window.scrollY > 1) {
			setScrolled(true);
		} else {
			setScrolled(false);
		}
	};

	return (
		<Box
			position={navbarPosition}
			boxShadow={navbarShadow}
			bg={navbarBg}
			borderColor={navbarBorder}
			filter={navbarFilter}
			backdropFilter={navbarBackdrop}
			backgroundPosition='center'
			backgroundSize='cover'
			borderRadius='16px'
			borderWidth='1.5px'
			borderStyle='solid'
			transitionDelay='0s, 0s, 0s, 0s'
			transitionDuration=' 0.25s, 0.25s, 0.25s, 0s'
			transition-property='box-shadow, background-color, filter, border'
			transitionTimingFunction='linear, linear, linear, linear'
			alignItems={{ xl: 'center' }}
			display={secondary ? 'block' : 'flex'}
			minH='75px'
			justifyContent={{ xl: 'center' }}
			lineHeight='25.6px'
			mx='auto'
			mt={secondaryMargin}
			pb='8px'
			right={{ base: '12px', md: '30px', lg: '30px', xl: '30px' }}
			px={{
				sm: paddingX,
				md: '10px'
			}}
			ps={{
				xl: '12px'
			}}
			pt='8px'
			top={{ base: '12px', md: '16px', lg: '20px', xl: '20px' }}
			w={{
				base: 'calc(100vw - 6%)',
				md: 'calc(100vw - 8%)',
				lg: 'calc(100vw - 6%)',
				xl: 'calc(100vw - 350px)',
				'2xl': 'calc(100vw - 365px)'
			}}>
			<Flex
				w='100%'
				flexDirection={{
					sm: 'column',
					md: 'row'
				}}
				alignItems={{ xl: 'center' }}
				mb={gap}>
				<Box mb={{ sm: '8px', md: '0px' }}>
					<Breadcrumb>
						<BreadcrumbItem color={secondaryText} fontSize='sm' mb='5px'>
							<BreadcrumbLink href='#' color={secondaryText}>
								Pages
							</BreadcrumbLink>
						</BreadcrumbItem>

						<BreadcrumbItem color={secondaryText} fontSize='sm' mb='5px'>
							<BreadcrumbLink href='#' color={secondaryText}>
								{brandText}
							</BreadcrumbLink>
						</BreadcrumbItem>
					</Breadcrumb>
					{/* Here we create navbar brand, based on route name */}
					<Link
						color={mainText}
						href='#'
						bg='inherit'
						borderRadius='inherit'
						fontWeight='bold'
						fontSize='34px'
						_hover={{ color: { mainText } }}
						_active={{
							bg: 'inherit',
							transform: 'none',
							borderColor: 'transparent'
						}}
						_focus={{
							boxShadow: 'none'
						}}>
						{brandText}
					</Link>
				</Box>
				<Box>
				<Flex alignItems="center" p="10px" boxShadow={scrolled ? 'md' : 'none'}>
				<Button colorScheme="teal" onClick={onOpen} ml="10px" borderRadius="5">
  Upload file
</Button>

	  {/* Add the FileManager button */}
	
</Flex>


			{/* Upload Modal */}
			<Modal isOpen={isOpen} onClose={onClose}>
				<ModalOverlay />
				<ModalContent>
					<ModalHeader>Upload Excel Files</ModalHeader>
					<ModalCloseButton />
					<ModalBody>
						<Input type="file" multiple onChange={handleFileChange} />
						<Button
							mt="10px"
							colorScheme="blue"
							isLoading={isValidating}
							onClick={validateHeaders}
						>
							Validate Files
						</Button>
					</ModalBody>
					<ModalFooter>
						<Button
							colorScheme="green"
							isDisabled={!validatedFiles.length}
							onClick={() => uploadFiles('lubricants')}
						>
							Save to Lubricants
						</Button>
						<Button
							colorScheme="orange"
							ml="10px"
							isDisabled={!validatedFiles.length}
							onClick={() => uploadFiles('petroleum')}
						>
							Save to Petroleum
						</Button>
					</ModalFooter>
				</ModalContent>
			</Modal>
		</Box>
				<Box ms='auto' w={{ sm: '100%', md: 'unset' }}>
					<AdminNavbarLinks
						onOpen={props.onOpen}
						logoText={props.logoText}
						secondary={props.secondary}
						fixed={props.fixed}
						scrolled={scrolled}
					/>
				</Box>
			</Flex>
			{secondary ? <Text color='white'>{message}</Text> : null}
			
		</Box>
	
	
	);
}

AdminNavbar.propTypes = {
	brandText: PropTypes.string,
	variant: PropTypes.string,
	secondary: PropTypes.bool,
	fixed: PropTypes.bool,
	onOpen: PropTypes.func
};
