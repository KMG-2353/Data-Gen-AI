import { useState } from "react";
import TextareaAutoGrowDemo from "./components/shadcn-studio/textarea/textarea-17.tsx";
import { Button } from "./components/ui/button.tsx";
import { Input } from "./components/ui/input";
import { MultiSelectSearch } from "./components/ui/multi-select-search";
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

const AUTO_POLICY_LOBS = ["Personal Auto Policy"];

const US_STATES: { value: string; label: string }[] = [
	{ value: "AL", label: "Alabama" },
	{ value: "AK", label: "Alaska" },
	{ value: "AZ", label: "Arizona" },
	{ value: "AR", label: "Arkansas" },
	{ value: "CA", label: "California" },
	{ value: "CO", label: "Colorado" },
	{ value: "CT", label: "Connecticut" },
	{ value: "DE", label: "Delaware" },
	{ value: "FL", label: "Florida" },
	{ value: "GA", label: "Georgia" },
	{ value: "HI", label: "Hawaii" },
	{ value: "ID", label: "Idaho" },
	{ value: "IL", label: "Illinois" },
	{ value: "IN", label: "Indiana" },
	{ value: "IA", label: "Iowa" },
	{ value: "KS", label: "Kansas" },
	{ value: "KY", label: "Kentucky" },
	{ value: "LA", label: "Louisiana" },
	{ value: "ME", label: "Maine" },
	{ value: "MD", label: "Maryland" },
	{ value: "MA", label: "Massachusetts" },
	{ value: "MI", label: "Michigan" },
	{ value: "MN", label: "Minnesota" },
	{ value: "MS", label: "Mississippi" },
	{ value: "MO", label: "Missouri" },
	{ value: "MT", label: "Montana" },
	{ value: "NE", label: "Nebraska" },
	{ value: "NV", label: "Nevada" },
	{ value: "NH", label: "New Hampshire" },
	{ value: "NJ", label: "New Jersey" },
	{ value: "NM", label: "New Mexico" },
	{ value: "NY", label: "New York" },
	{ value: "NC", label: "North Carolina" },
	{ value: "ND", label: "North Dakota" },
	{ value: "OH", label: "Ohio" },
	{ value: "OK", label: "Oklahoma" },
	{ value: "OR", label: "Oregon" },
	{ value: "PA", label: "Pennsylvania" },
	{ value: "RI", label: "Rhode Island" },
	{ value: "SC", label: "South Carolina" },
	{ value: "SD", label: "South Dakota" },
	{ value: "TN", label: "Tennessee" },
	{ value: "TX", label: "Texas" },
	{ value: "UT", label: "Utah" },
	{ value: "VT", label: "Vermont" },
	{ value: "VA", label: "Virginia" },
	{ value: "WA", label: "Washington" },
	{ value: "WV", label: "West Virginia" },
	{ value: "WI", label: "Wisconsin" },
	{ value: "WY", label: "Wyoming" },
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
	const [driverCount, setDriverCount] = useState<number | "">("");
	const [vehicleCount, setVehicleCount] = useState<number | "">("");
	const [generatedSheets, setGeneratedSheets] = useState<string[] | null>(null);

	const isAutoPolicy = selectedLobs.some((lob) =>
		AUTO_POLICY_LOBS.includes(lob),
	);

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
					...(isAutoPolicy && driverCount && { driver_count: driverCount }),
					...(isAutoPolicy && vehicleCount && { vehicle_count: vehicleCount }),
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


	const lobOptions =
		lineOfBusiness === "Monoline"
			? MONOLINE_LOBS
			: lineOfBusiness === "Multiline"
				? MULTILINE_LOBS
				: [];

	return (
		<div className="flex min-h-screen m-auto w-full max-w-2xl justify-center items-center flex-col gap-5 px-4 py-8 sm:px-6 md:px-8">
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
					<SelectTrigger className="w-full sm:max-w-48">
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
					<MultiSelectSearch
						options={US_STATES}
						selected={selectedStates}
						onChange={setSelectedStates}
						placeholder="Select states..."
						searchPlaceholder="Search by abbreviation or name..."
					/>
				</div>

				{/* Driver & Vehicle counts — only for auto policies */}
				{isAutoPolicy && (
					<div className="w-full flex gap-3">
						<div className="flex flex-col gap-1 flex-1">
							<p className="text-xs text-gray-500">Number of Drivers (optional)</p>
							<Input
								type="number"
								min={1}
								placeholder="e.g. 2"
								value={driverCount}
								onChange={(e) => setDriverCount(e.target.value === "" ? "" : Number(e.target.value))}
							/>
						</div>
						<div className="flex flex-col gap-1 flex-1">
							<p className="text-xs text-gray-500">Number of Vehicles (optional)</p>
							<Input
								type="number"
								min={1}
								placeholder="e.g. 3"
								value={vehicleCount}
								onChange={(e) => setVehicleCount(e.target.value === "" ? "" : Number(e.target.value))}
							/>
						</div>
					</div>
				)}

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
