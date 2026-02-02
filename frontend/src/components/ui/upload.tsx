import { useCallback, useState } from "react";
import { useDropzone } from "react-dropzone";
import { FaCloudUploadAlt } from "react-icons/fa";

interface FileUploadProps {
	onFileSelect?: (file: File) => void;
	accept?: string;
	maxSize?: number;
}

export default function FileUpload({
	onFileSelect,
	accept = ".xlsx,.xls",
	maxSize = 10 * 1024 * 1024, // 10MB default
}: FileUploadProps = {}) {
	const [selectedFile, setSelectedFile] = useState<File | null>(null);

	const onDrop = useCallback(
		(acceptedFiles: File[]) => {
			if (acceptedFiles && acceptedFiles.length > 0) {
				const file = acceptedFiles[0];
				setSelectedFile(file);
				onFileSelect?.(file);
			}
		},
		[onFileSelect],
	);

	const { getRootProps, getInputProps, isDragActive } = useDropzone({
		onDrop,
		accept: {
			"application/vnd.ms-excel": [".xls"],
			"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [
				".xlsx",
			],
		},
		maxSize,
		multiple: false,
	});

	return (
		<div
			{...getRootProps()}
			className={`relative border-2 border-dashed rounded-lg p-8 transition-colors cursor-pointer ${
				isDragActive
					? "border-blue-500 bg-blue-50"
					: "border-gray-300 hover:border-blue-400"
			} ${selectedFile ? "border-green-500 bg-green-50" : ""}`}
		>
			<input {...getInputProps()} />
			<div className="flex flex-col items-center justify-center gap-4">
				<FaCloudUploadAlt />
				{selectedFile ? (
					<div className="text-center">
						<p className="text-sm font-medium text-gray-700">
							{selectedFile.name}
						</p>
						<p className="text-xs text-gray-500">
							{(selectedFile.size / 1024).toFixed(2)} KB
						</p>
					</div>
				) : (
					<div className="text-center">
						<p className="text-sm font-medium text-gray-700">
							{isDragActive
								? "Drop the file here"
								: "Drag & drop or click to upload"}
						</p>
						<p className="text-xs text-gray-500 mt-1">
							Excel files (.xlsx, .xls) up to {maxSize / (1024 * 1024)}MB
						</p>
					</div>
				)}
			</div>
		</div>
	);
}
