import type { AnalysisProgress as AnalysisProgressType } from "@/types";

interface AnalysisProgressProps {
	progress: AnalysisProgressType;
}

export function AnalysisProgress({ progress }: AnalysisProgressProps) {
	if (progress.status === "idle") {
		return null;
	}

	return (
		<div className="w-full p-4 bg-blue-50 rounded-lg border border-blue-100">
			<div className="flex items-center justify-between mb-2">
				<span className="text-sm font-medium text-blue-800">
					{progress.status === "analyzing" && "Analyzing patterns..."}
					{progress.status === "complete" && "Analysis complete!"}
					{progress.status === "error" && "Analysis error"}
				</span>
				<span className="text-sm text-blue-600">
					{Math.round(progress.progress)}%
				</span>
			</div>

			{/* Progress bar */}
			<div className="w-full h-2 bg-blue-100 rounded-full overflow-hidden mb-2">
				<div
					className={`h-full transition-all duration-300 rounded-full ${
						progress.status === "error" ? "bg-red-500" : "bg-blue-500"
					}`}
					style={{ width: `${progress.progress}%` }}
				/>
			</div>

			{/* Current sheet indicator */}
			{progress.current_sheet && progress.status === "analyzing" && (
				<div className="flex items-center gap-2">
					<div className="w-2 h-2 bg-blue-500 rounded-full animate-pulse" />
					<span className="text-xs text-blue-700">
						Processing: {progress.current_sheet}
					</span>
				</div>
			)}

			{/* Completed sheets */}
			{progress.sheets_completed.length > 0 && (
				<div className="mt-2 flex flex-wrap gap-1">
					{progress.sheets_completed.map((sheet) => (
						<span
							key={sheet}
							className="px-2 py-0.5 bg-green-100 text-green-700 text-xs rounded-full"
						>
							{sheet}
						</span>
					))}
				</div>
			)}

			{/* Error message */}
			{progress.status === "error" && progress.error_message && (
				<p className="mt-2 text-sm text-red-600">{progress.error_message}</p>
			)}
		</div>
	);
}
