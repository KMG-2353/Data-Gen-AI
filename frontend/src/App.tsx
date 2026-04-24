import { useState } from "react";
import TextareaAutoGrowDemo from "./components/shadcn-studio/textarea/textarea-17.tsx";
import { Button } from "./components/ui/button.tsx";
import { Input } from "./components/ui/input";
import {
	Select,
	SelectContent,
	SelectGroup,
	SelectItem,
	SelectLabel,
	SelectTrigger,
	SelectValue,
} from "./components/ui/select";
import FileUpload from "./components/ui/upload.tsx";
import { apiConfig } from "./config/api";

const MULTILINE_LOBS = [
	"Inland Marine",
	"Crime",
	// "General Liability",
	// "Optional Coverage",
	// "Commercial Auto",
	"Property",
];
const MONOLINE_LOBS = ["GL", "Personal Auto Policy"];
const US_STATES = [
	"CA",
	"TX",
	"FL",
	"NY",
	"PA",
	"IL",
	"OH",
	"GA",
	"NC",
	"MI",
	"ME",
	"MA",
];

function App() {
	const [headersBySheet, setHeadersBySheet] = useState<
		Record<string, string[]>
	>({});
	const [sheetNames, setSheetNames] = useState<string[]>([]);
	const [sessionId, setSessionId] = useState<string | null>(null);
	const [sepcialinstruction, setSepcialinstruction] = useState("");
	const [isUploading, setIsUploading] = useState(false);
	const [isGenerating, setIsGenerating] = useState(false);
	const [lineOfBusiness, setLineOfBusiness] = useState<string>("");
	const [selectedLobs, setSelectedLobs] = useState<string[]>([]);
	const [selectedStates, setSelectedStates] = useState<string[]>([]);
	const [testCases, setTestCases] = useState(5);
	const [generatedSheets, setGeneratedSheets] = useState<string[] | null>(null);

	const handleFileSelect = async (file: File) => {
		setIsUploading(true);
		setGeneratedSheets(null);

		try {
			const formData = new FormData();
			formData.append("file", file);

			const response = await fetch(apiConfig.endpoints.upload, {
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
			const response = await fetch(apiConfig.endpoints.generate, {
				method: "POST",
				headers: { "Content-Type": "application/json" },
				body: JSON.stringify({
					session_id: sessionId,
					row_count: testCases,
					special_inst: sepcialinstruction,
					lob_type: lineOfBusiness,
					lob_selection: selectedLobs,
					state_selection: selectedStates,
				}),
			});

			if (!response.ok) throw new Error("Generation failed");

			const data = await response.json();
			setGeneratedSheets(data.sheets_generated);
		} catch (error) {
			console.error("Generation error:", error);
			alert("Failed to generate test cases. Please try again.");
		} finally {
			setIsGenerating(false);
		}
	};

	const handleDownload = async () => {
		if (!sessionId) return;

		const response = await fetch(apiConfig.endpoints.download(sessionId));
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

	const toggleLob = (lob: string) => {
		if (lineOfBusiness === "Monoline") {
			setSelectedLobs([lob]);
		} else {
			setSelectedLobs((prev) =>
				prev.includes(lob) ? prev.filter((l) => l !== lob) : [...prev, lob],
			);
		}
	};

	const toggleState = (state: string) => {
		setSelectedStates((prev) =>
			prev.includes(state) ? prev.filter((s) => s !== state) : [...prev, state],
		);
	};

	const lobOptions =
		lineOfBusiness === "Monoline"
			? MONOLINE_LOBS
			: lineOfBusiness === "Multiline"
				? MULTILINE_LOBS
				: [];

	return (
		<div className="flex h-screen m-auto w-1/2 max-w-2xl justify-center items-center flex-col gap-5 p-8">
			<h1 className="text-2xl font-bold">Data Gen Agent</h1>
			<div className="w-full flex flex-col gap-3">
				<h2 className="text-sm font-medium">
					Upload your sample data sheet (accepted format - .xlsx)
				</h2>
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
				<h2 className="text-sm font-medium">
					Number of test data records to be generated{" "}
				</h2>
				<Input
					type="number"
					value={testCases}
					onChange={(e) => setTestCases(Number(e.target.value))}
					min={1}
					max={500}
				/>

				{/* LOB type selector */}
				<Select
					value={lineOfBusiness}
					onValueChange={(val) => {
						setLineOfBusiness(val);
						setSelectedLobs([]);
					}}
				>
					<SelectTrigger className="w-full max-w-48">
						<SelectValue placeholder="Select LOB Type" />
					</SelectTrigger>
					<SelectContent>
						<SelectGroup>
							<SelectLabel>Line of Business</SelectLabel>
							<SelectItem value="Monoline">Monoline</SelectItem>
							<SelectItem value="Multiline">Multiline</SelectItem>
						</SelectGroup>
					</SelectContent>
				</Select>

				{/* LOB toggle buttons */}
				{lobOptions.length > 0 && (
					<div className="w-full flex flex-col gap-2">
						<p className="text-xs text-gray-500">
							{lineOfBusiness === "Monoline"
								? "Select one LOB"
								: "Select one or more LOBs (optional — defaults to random if none selected)"}
						</p>
						<div className="flex flex-wrap gap-2">
							{lobOptions.map((lob) => {
								const active = selectedLobs.includes(lob);
								return (
									<button
										key={lob}
										type="button"
										onClick={() => toggleLob(lob)}
										className={`px-3 py-1 rounded-full text-xs border transition-colors ${
											active
												? "bg-gray-800 text-white border-gray-800"
												: "bg-white text-gray-700 border-gray-300 hover:border-gray-500"
										}`}
									>
										{lob}
									</button>
								);
							})}
						</div>
					</div>
				)}

				{/* State selector */}
				<div className="w-full flex flex-col gap-2">
					<p className="text-xs text-gray-500">
						Select states (optional — defaults to diverse if none selected)
					</p>
					<div className="flex flex-wrap gap-2">
						{US_STATES.map((state) => {
							const active = selectedStates.includes(state);
							return (
								<button
									key={state}
									type="button"
									onClick={() => toggleState(state)}
									className={`px-3 py-1 rounded-full text-xs border transition-colors ${
										active
											? "bg-gray-800 text-white border-gray-800"
											: "bg-white text-gray-700 border-gray-300 hover:border-gray-500"
									}`}
								>
									{state}
								</button>
							);
						})}
					</div>
				</div>

				<TextareaAutoGrowDemo
					value={sepcialinstruction}
					inputchange={setSepcialinstruction}
				/>
			</div>

			<Button
				onClick={handleGenerateTestCases}
				disabled={!sessionId || isUploading || isGenerating}
				className="w-full"
			>
				{isGenerating ? "Generating..." : "Generate Test Data"}
			</Button>
			{generatedSheets && (
				<div className="w-full flex flex-col gap-3">
					<h2 className="text-sm font-medium">
						✅ Test Data Generated for {generatedSheets.length} sheet
						{generatedSheets.length > 1 ? "s" : ""}
					</h2>
					<div className="border rounded-lg p-4">
						<ul className="text-sm text-gray-600 space-y-1">
							{generatedSheets.map((sheetName) => (
								<li key={sheetName}>• {sheetName}</li>
							))}
						</ul>
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
