import { useState } from "react";
import type { ColumnRule, SheetRuleSet } from "@/types";
import { RuleCard } from "./RuleCard";

interface RuleSetPanelProps {
	ruleSets: Record<string, SheetRuleSet>;
	sheetNames: string[];
	onUpdateRule: (
		sheetName: string,
		columnName: string,
		rule: ColumnRule,
	) => Promise<void>;
	onRepromptRule: (
		sheetName: string,
		columnName: string,
		feedback: string,
	) => Promise<void>;
}

export function RuleSetPanel({
	ruleSets,
	sheetNames,
	onUpdateRule,
	onRepromptRule,
}: RuleSetPanelProps) {
	const [activeSheet, setActiveSheet] = useState<string>(sheetNames[0] || "");

	const activeRuleSet = ruleSets[activeSheet];

	return (
		<div className="w-full">
			{/* Sheet tabs */}
			{sheetNames.length > 1 && (
				<div className="flex gap-1 mb-4 overflow-x-auto pb-2">
					{sheetNames.map((sheetName) => (
						<button
							key={sheetName}
							onClick={() => setActiveSheet(sheetName)}
							className={`px-3 py-1.5 text-sm rounded-md whitespace-nowrap transition-colors ${
								activeSheet === sheetName
									? "bg-primary text-primary-foreground"
									: "bg-gray-100 text-gray-700 hover:bg-gray-200"
							}`}
						>
							{sheetName}
							{ruleSets[sheetName] && (
								<span className="ml-1.5 text-xs opacity-70">
									({ruleSets[sheetName].rules.length})
								</span>
							)}
						</button>
					))}
				</div>
			)}

			{/* Rules list */}
			{activeRuleSet ? (
				<div className="space-y-3">
					{activeRuleSet.rules.map((rule) => (
						<RuleCard
							key={rule.column_name}
							rule={rule}
							sheetName={activeSheet}
							onUpdate={(updatedRule) =>
								onUpdateRule(activeSheet, rule.column_name, updatedRule)
							}
							onReprompt={(feedback) =>
								onRepromptRule(activeSheet, rule.column_name, feedback)
							}
						/>
					))}

					{/* Cross-column rules */}
					{activeRuleSet.cross_column_rules &&
						activeRuleSet.cross_column_rules.length > 0 && (
							<div className="mt-4 p-3 bg-gray-50 rounded-lg">
								<h4 className="text-sm font-medium mb-2 text-gray-700">
									Cross-Column Rules
								</h4>
								<ul className="text-sm text-gray-600 space-y-1">
									{activeRuleSet.cross_column_rules.map((rule, idx) => (
										<li key={idx} className="flex items-start gap-2">
											<span className="text-gray-400">•</span>
											{rule}
										</li>
									))}
								</ul>
							</div>
						)}
				</div>
			) : (
				<div className="text-center py-8 text-gray-500">
					<p>No rules analyzed yet</p>
					<p className="text-sm mt-1">
						Click "Analyze Patterns" to learn rules from your data
					</p>
				</div>
			)}
		</div>
	);
}
