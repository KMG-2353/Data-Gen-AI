import { useState } from "react";
import { Button } from "./components/ui/button.tsx";
import { Input } from "./components/ui/input";
import FileUpload from "./components/ui/upload.tsx";

function App() {
	const [selectedFile, setSelectedFile] = useState<File | null>(null);
	const [headers, setHeaders] = useState<string[]>([]);
	const [sessionId, setSessionId] = useState<string | null>(null);
	const [isUploading, setIsUploading] = useState(false);
	const [testCases, setTestCases] = useState(10);

	const handleFileSelect = async (file: File) => {
		setSelectedFile(file);
		setIsUploading(true);

		try {
			const formData = new FormData();
			formData.append("file", file);

			const response = await fetch("/api/upload", {
				method: "POST",
				body: formData,
			});

			if (!response.ok) {
				throw new Error("Upload failed");
			}

			const data = await response.json();
			setHeaders(data.headers);
			setSessionId(data.session_id);
		} catch (error) {
			console.error("Upload error:", error);
			alert("Failed to upload file. Please try again.");
		} finally {
			setIsUploading(false);
		}
	};

	const handleGenerateTestCases = () => {
		if (!sessionId) {
			alert("Please upload an Excel file first");
			return;
		}
		console.log("Generating", testCases, "test cases for session:", sessionId);
		// TODO: Implement test case generation
	};

	return (
		<div className="flex h-screen m-auto w-1/2 max-w-2xl justify-center items-center flex-col gap-5">
			<h1 className="text-2xl font-bold">Testing Agent</h1>

			<div className="w-full flex flex-col gap-3">
				<h2 className="text-sm font-medium">Upload Excel File</h2>
				<FileUpload onFileSelect={handleFileSelect} />
				{isUploading && <p className="text-sm text-gray-500">Uploading...</p>}
			</div>

			{headers.length > 0 && (
				<div className="w-full flex flex-col gap-3">
					<h2 className="text-sm font-medium">
						Extracted Headers ({headers.length})
					</h2>
					<div className="border rounded-lg p-4 max-h-40 overflow-y-auto">
						<div className="flex flex-wrap gap-2">
							{headers.map((header, index) => (
								<span
									key={index}
									className="px-3 py-1 bg-blue-100 text-blue-800 rounded-full text-xs"
								>
									{header}
								</span>
							))}
						</div>
					</div>
				</div>
			)}

			<div className="w-full flex flex-col gap-3">
				<h2 className="text-sm font-medium">Number of Test Cases</h2>
				<Input
					type="number"
					placeholder="Enter number of test cases"
					value={testCases}
					onChange={(e) => setTestCases(Number(e.target.value))}
					min={1}
					max={500}
				/>
			</div>

			<Button
				onClick={handleGenerateTestCases}
				disabled={!sessionId || isUploading}
				className="w-full"
			>
				Generate Test Cases
			</Button>
		</div>
	);
}
export default App;
