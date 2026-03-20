import { useState, useEffect, useCallback } from "react";
import TextareaAutoGrowDemo from "./components/shadcn-studio/textarea/textarea-17.tsx";
import { Button } from "./components/ui/button.tsx";
import { Input } from "./components/ui/input";
import { Switch } from "./components/ui/switch";
import { Label } from "./components/ui/label";
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
import { RuleSetPanel } from "./components/rules/RuleSetPanel";
import { AnalysisProgress } from "./components/rules/AnalysisProgress";
import { apiConfig, warmUpServer } from "./config/api";
import type {
	SheetData,
	SheetRuleSet,
	ColumnRule,
	AnalysisProgress as AnalysisProgressType,
	AnalysisEvent,
} from "./types";

function App() {
	// Upload state
	const [sheetsData, setSheetsData] = useState<SheetData[]>([]);
	const [sheetNames, setSheetNames] = useState<string[]>([]);
	const [sessionId, setSessionId] = useState<string | null>(null);
	const [isUploading, setIsUploading] = useState(false);

	// Rules state
	const [verifyRulesBeforeGeneration, setVerifyRulesBeforeGeneration] =
		useState(true);
	const [ruleSets, setRuleSets] = useState<Record<string, SheetRuleSet>>({});
	const [analysisProgress, setAnalysisProgress] =
		useState<AnalysisProgressType>({
			current_sheet: null,
			sheets_completed: [],
			sheets_total: 0,
			progress: 0,
			status: "idle",
		});

	// Generation state
	const [specialInstruction, setSpecialInstruction] = useState("");
	const [isGenerating, setIsGenerating] = useState(false);
	const [lineOfBusiness, setLineOfBusiness] = useState<string>("");
	const [coverage, setCoverage] = useState<string>("");
	const [testCases, setTestCases] = useState(5);
	const [generatedSheets, setGeneratedSheets] = useState<string[] | null>(null);

	// Warm up server on mount
	useEffect(() => {
		warmUpServer();
	}, []);

	const handleFileSelect = async (file: File) => {
		setIsUploading(true);
		setGeneratedSheets(null);
		setRuleSets({});
		setAnalysisProgress({
			current_sheet: null,
			sheets_completed: [],
			sheets_total: 0,
			progress: 0,
			status: "idle",
		});

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
			setSheetsData(data.sheets || []);
			setSheetNames(data.sheet_names);
			setSessionId(data.session_id);
		} catch (error) {
			console.error("Upload error:", error);
			alert("Failed to upload file. Please try again.");
		} finally {
			setIsUploading(false);
		}
	};

	const handleAnalyzePatterns = useCallback(async () => {
		if (!sessionId) return;

		setAnalysisProgress({
			current_sheet: null,
			sheets_completed: [],
			sheets_total: sheetNames.length,
			progress: 0,
			status: "analyzing",
		});

		try {
			const response = await fetch(apiConfig.endpoints.analyze, {
				method: "POST",
				headers: { "Content-Type": "application/json" },
				body: JSON.stringify({ session_id: sessionId }),
			});

			if (!response.ok) throw new Error("Analysis failed");

			const reader = response.body?.getReader();
			const decoder = new TextDecoder();

			if (!reader) throw new Error("No response body");

			while (true) {
				const { done, value } = await reader.read();
				if (done) break;

				const text = decoder.decode(value);
				const lines = text.split("\n");

				for (const line of lines) {
					if (line.startsWith("data: ")) {
						try {
							const event: AnalysisEvent = JSON.parse(line.slice(6));

							switch (event.event) {
								case "sheet_start":
									setAnalysisProgress((prev) => ({
										...prev,
										current_sheet: event.sheet_name || null,
										progress: event.progress * 100,
									}));
									break;

								case "sheet_complete":
									if (event.sheet_name && event.rules) {
										setRuleSets((prev) => ({
											...prev,
											[event.sheet_name!]: event.rules!,
										}));
									}
									setAnalysisProgress((prev) => ({
										...prev,
										sheets_completed: event.sheet_name
											? [...prev.sheets_completed, event.sheet_name]
											: prev.sheets_completed,
										progress: event.progress * 100,
									}));
									break;

								case "complete":
									setAnalysisProgress((prev) => ({
										...prev,
										status: "complete",
										current_sheet: null,
										progress: 100,
									}));
									break;

								case "error":
									setAnalysisProgress((prev) => ({
										...prev,
										status: "error",
										error_message: event.message,
									}));
									break;
							}
						} catch (e) {
							console.error("Error parsing SSE event:", e);
						}
					}
				}
			}
		} catch (error) {
			console.error("Analysis error:", error);
			setAnalysisProgress((prev) => ({
				...prev,
				status: "error",
				error_message: "Failed to analyze patterns",
			}));
		}
	}, [sessionId, sheetNames.length]);

	const handleUpdateRule = async (
		sheetName: string,
		columnName: string,
		updatedRule: ColumnRule
	) => {
		if (!sessionId) return;

		const response = await fetch(apiConfig.endpoints.updateRule, {
			method: "PUT",
			headers: { "Content-Type": "application/json" },
			body: JSON.stringify({
				session_id: sessionId,
				sheet_name: sheetName,
				column_name: columnName,
				updated_rule: updatedRule,
			}),
		});

		if (!response.ok) throw new Error("Update failed");

		const data = await response.json();

		// Update local state
		setRuleSets((prev) => {
			const sheetRules = prev[sheetName];
			if (!sheetRules) return prev;

			return {
				...prev,
				[sheetName]: {
					...sheetRules,
					rules: sheetRules.rules.map((r) =>
						r.column_name === columnName ? data.updated_rule : r
					),
				},
			};
		});
	};

	const handleRepromptRule = async (
		sheetName: string,
		columnName: string,
		feedback: string
	) => {
		if (!sessionId) return;

		const response = await fetch(apiConfig.endpoints.repromptRule, {
			method: "POST",
			headers: { "Content-Type": "application/json" },
			body: JSON.stringify({
				session_id: sessionId,
				sheet_name: sheetName,
				column_name: columnName,
				user_feedback: feedback,
			}),
		});

		if (!response.ok) throw new Error("Reprompt failed");

		const data = await response.json();

		// Update local state
		setRuleSets((prev) => {
			const sheetRules = prev[sheetName];
			if (!sheetRules) return prev;

			return {
				...prev,
				[sheetName]: {
					...sheetRules,
					rules: sheetRules.rules.map((r) =>
						r.column_name === columnName ? data.new_rule : r
					),
				},
			};
		});
	};

	const handleGenerateTestCases = async () => {
		if (!sessionId) {
			alert("Please upload an Excel file first");
			return;
		}

		// If verify mode is on and no rules analyzed, analyze first
		if (
			verifyRulesBeforeGeneration &&
			Object.keys(ruleSets).length === 0 &&
			analysisProgress.status !== "complete"
		) {
			alert("Please analyze patterns first before generating");
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
					special_inst: specialInstruction,
					skip_rules: !verifyRulesBeforeGeneration,
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

	const hasRules = Object.keys(ruleSets).length > 0;

	return (
		<div className="flex min-h-screen m-auto w-full max-w-3xl justify-start items-center flex-col gap-5 p-8">
			<h1 className="text-2xl font-bold">Data Gen Agent V2</h1>

			{/* Verify Rules Toggle */}
			<div className="w-full flex items-center justify-between p-3 bg-gray-50 rounded-lg">
				<div className="flex flex-col">
					<Label htmlFor="verify-toggle" className="text-sm font-medium">
						Verify rules before generation
					</Label>
					<span className="text-xs text-gray-500">
						Analyze patterns and review rules before generating data
					</span>
				</div>
				<Switch
					id="verify-toggle"
					checked={verifyRulesBeforeGeneration}
					onCheckedChange={setVerifyRulesBeforeGeneration}
				/>
			</div>

			{/* File Upload */}
			<div className="w-full flex flex-col gap-3">
				<h2 className="text-sm font-medium">
					Upload your sample data sheet (accepted format - .xlsx)
				</h2>
				<FileUpload onFileSelect={handleFileSelect} />
				{isUploading && <p className="text-sm text-gray-500">Uploading...</p>}
			</div>

			{/* Extracted Headers */}
			{sheetsData.length > 0 && (
				<div className="w-full flex flex-col gap-3">
					<h2 className="text-sm font-medium">
						Extracted Data ({sheetNames.length} sheet
						{sheetNames.length > 1 ? "s" : ""})
					</h2>
					<div className="border rounded-lg p-4 max-h-48 overflow-y-auto space-y-4">
						{sheetsData.map((sheet) => (
							<div key={sheet.sheet_name}>
								<div className="flex items-center justify-between mb-2">
									<p className="text-xs font-semibold text-gray-600">
										{sheet.sheet_name}
									</p>
									<span className="text-xs text-gray-400">
										{sheet.total_rows} rows, {sheet.sample_count} samples
									</span>
								</div>
								<div className="flex flex-wrap gap-2">
									{sheet.headers?.map((header, index) => (
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

			{/* Analyze Patterns Button */}
			{verifyRulesBeforeGeneration && sessionId && !hasRules && (
				<Button
					onClick={handleAnalyzePatterns}
					disabled={
						analysisProgress.status === "analyzing" || isUploading || hasRules
					}
					className="w-full"
					variant="outline"
				>
					{analysisProgress.status === "analyzing"
						? "Analyzing..."
						: "Analyze Patterns"}
				</Button>
			)}

			{/* Analysis Progress */}
			{analysisProgress.status !== "idle" && (
				<AnalysisProgress progress={analysisProgress} />
			)}

			{/* Rules Panel */}
			{verifyRulesBeforeGeneration && hasRules && (
				<div className="w-full flex flex-col gap-3">
					<h2 className="text-sm font-medium">
						Learned Rules ({Object.keys(ruleSets).length} sheet
						{Object.keys(ruleSets).length > 1 ? "s" : ""})
					</h2>
					<div className="border rounded-lg p-4 max-h-96 overflow-y-auto">
						<RuleSetPanel
							ruleSets={ruleSets}
							sheetNames={sheetNames}
							onUpdateRule={handleUpdateRule}
							onRepromptRule={handleRepromptRule}
						/>
					</div>
				</div>
			)}

			{/* Generation Config */}
			<div className="w-full flex flex-col gap-3">
				<h2 className="text-sm font-medium">
					Number of test data records to be generated
				</h2>
				<Input
					type="number"
					value={testCases}
					onChange={(e) => setTestCases(Number(e.target.value))}
					min={1}
					max={500}
				/>
				<div className="w-full flex gap-3">
					<Select
						value={lineOfBusiness}
						onValueChange={(val) => {
							setLineOfBusiness(val);
							setCoverage("");
						}}
					>
						<SelectTrigger className="w-full max-w-48">
							<SelectValue placeholder="Select LOB" />
						</SelectTrigger>
						<SelectContent>
							<SelectGroup>
								<SelectLabel>Select LOB</SelectLabel>
								<SelectItem value="Monoline">Monoline</SelectItem>
								<SelectItem value="Multiline">Multiline</SelectItem>
							</SelectGroup>
						</SelectContent>
					</Select>

					{lineOfBusiness && (
						<Select value={coverage} onValueChange={setCoverage}>
							{lineOfBusiness === "Monoline" ? (
								<SelectTrigger className="w-full max-w-56">
									<SelectValue placeholder="Select" />
								</SelectTrigger>
							) : (
								<SelectTrigger className="w-full max-w-56">
									<SelectValue placeholder="Select Multiple" />
								</SelectTrigger>
							)}

							<SelectContent>
								<SelectGroup>
									{lineOfBusiness === "Monoline" ? (
										<div>
											<SelectLabel>Select</SelectLabel>
											<SelectItem value="GL">GL</SelectItem>
										</div>
									) : (
										<>
											<SelectLabel>Select Multi</SelectLabel>
											<SelectItem value="Inline Marine">
												Inline Marine
											</SelectItem>
											<SelectItem value="Crime">Crime</SelectItem>
											<SelectItem value="General Liability">
												General Liability
											</SelectItem>
											<SelectItem value="Optional Coverage">
												Optional Coverage
											</SelectItem>
											<SelectItem value="Commercial Auto">
												Commercial Auto
											</SelectItem>
										</>
									)}
								</SelectGroup>
							</SelectContent>
						</Select>
					)}
				</div>
				<TextareaAutoGrowDemo
					value={specialInstruction}
					inputchange={setSpecialInstruction}
				/>
			</div>

			{/* Generate Button */}
			<Button
				onClick={handleGenerateTestCases}
				disabled={
					!sessionId ||
					isUploading ||
					isGenerating ||
					(verifyRulesBeforeGeneration && !hasRules)
				}
				className="w-full"
			>
				{isGenerating ? "Generating..." : "Generate Test Data"}
			</Button>

			{/* Download Section */}
			{generatedSheets && (
				<div className="w-full flex flex-col gap-3">
					<h2 className="text-sm font-medium">
						Test Data Generated for {generatedSheets.length} sheet
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
