import { useState } from "react";
import { Button } from "./components/ui/button.tsx";
import { Input } from "./components/ui/input";
import FileUpload from "./components/ui/upload.tsx";

function App() {
	const [headersBySheet, setHeadersBySheet] = useState<
		Record<string, string[]>
	>({});
	const [sheetNames, setSheetNames] = useState<string[]>([]);
	const [sessionId, setSessionId] = useState<string | null>(null);
	const [isUploading, setIsUploading] = useState(false);
	const [isGenerating, setIsGenerating] = useState(false);
	const [testCases, setTestCases] = useState(10);
	const [generatedData, setGeneratedData] = useState<
		Record<string, unknown>[] | null
	>(null);

	const handleFileSelect = async (file: File) => {
		setIsUploading(true);
		setGeneratedData(null);

		try {
			const formData = new FormData();
			formData.append("file", file);

			const response = await fetch("/api/upload", {
				method: "POST",
				body: formData,
			});

			if (!response.ok) throw new Error("Upload failed");

			const data = await response.json();
			console.log("Upload response:", data);
			setHeadersBySheet(data.headers_by_sheet);
			setSheetNames(data.sheet_names);
			setSessionId(data.session_id);
		} catch (error) {
			console.error("Upload error:", error);
			alert("Failed to upload file. Please try again.");
		} finally {
			setIsUploading(false);
		}
	};

	const handleGenerateTestCases = async () => {
		if (!sessionId) {
			alert("Please upload an Excel file first");
			return;
		}

		setIsGenerating(true);
		try {
			const response = await fetch("/api/generate", {
				method: "POST",
				headers: { "Content-Type": "application/json" },
				body: JSON.stringify({
					session_id: sessionId,
					row_count: testCases,
				}),
			});

			if (!response.ok) throw new Error("Generation failed");

			const data = await response.json();
			setGeneratedData(data.data);
		} catch (error) {
			console.error("Generation error:", error);
			alert("Failed to generate test cases. Please try again.");
		} finally {
			setIsGenerating(false);
		}
	};

	const handleDownload = async () => {
		if (!sessionId) return;

		const response = await fetch(`/api/download/${sessionId}`);
		const blob = await response.blob();

		const url = URL.createObjectURL(blob);
		const a = document.createElement("a");
		a.href = url;
		a.download = `test_data_${sessionId.slice(0, 8)}.xlsx`;
		document.body.appendChild(a);
		a.click();
		document.body.removeChild(a);
		URL.revokeObjectURL(url);
	};

	return (
		<div className="flex h-screen m-auto w-1/2 max-w-2xl justify-center items-center flex-col gap-5 p-8">
			<h1 className="text-2xl font-bold">AI Testing Agent</h1>

			<div className="w-full flex flex-col gap-3">
				<h2 className="text-sm font-medium">Upload Excel File</h2>
				<FileUpload onFileSelect={handleFileSelect} />
				{isUploading && <p className="text-sm text-gray-500">Uploading...</p>}
			</div>

			{sheetNames.length > 0 && (
				<div className="w-full flex flex-col gap-3">
					<h2 className="text-sm font-medium">
						Extracted Headers ({sheetNames.length} sheet
						{sheetNames.length > 1 ? "s" : ""})
					</h2>
					<div className="border rounded-lg p-4 max-h-60 overflow-y-auto space-y-4">
						{sheetNames.map((sheetName) => (
							<div key={sheetName}>
								<p className="text-xs font-semibold text-gray-600 mb-2">
									{sheetName}
								</p>
								<div className="flex flex-wrap gap-2">
									{headersBySheet[sheetName]?.map((header, index) => (
										<span
											key={index}
											className="px-3 py-1 bg-blue-100 text-blue-800 rounded-full text-xs"
										>
											{header}
										</span>
									))}
								</div>
							</div>
						))}
					</div>
				</div>
			)}

			<div className="w-full flex flex-col gap-3">
				<h2 className="text-sm font-medium">Number of Test Cases</h2>
				<Input
					type="number"
					value={testCases}
					onChange={(e) => setTestCases(Number(e.target.value))}
					min={1}
					max={500}
				/>
			</div>

			<Button
				onClick={handleGenerateTestCases}
				disabled={!sessionId || isUploading || isGenerating}
				className="w-full"
			>
				{isGenerating ? "Generating..." : "Generate Test Cases"}
			</Button>

			{generatedData && (
				<div className="w-full flex flex-col gap-3">
					<h2 className="text-sm font-medium">
						Generated Data ({generatedData.length} rows)
					</h2>
					<div className="border rounded-lg p-4 max-h-60 overflow-auto">
						<pre className="text-xs">
							{JSON.stringify(generatedData.slice(0, 3), null, 2)}
						</pre>
						{generatedData.length > 3 && (
							<p className="text-xs text-gray-500 mt-2">
								...and {generatedData.length - 3} more rows
							</p>
						)}
					</div>
					<Button onClick={handleDownload} variant="outline" className="w-full">
						Download Excel
					</Button>
				</div>
			)}
		</div>
	);
}

export default App;
