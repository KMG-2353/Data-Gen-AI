import { useState } from "react";
import { Button } from "@/components/ui/button";
import { Textarea } from "@/components/ui/textarea";
import type { ColumnRule } from "@/types";

interface RuleCardProps {
	rule: ColumnRule;
	sheetName: string;
	onUpdate: (rule: ColumnRule) => Promise<void>;
	onReprompt: (feedback: string) => Promise<void>;
}

export function RuleCard({ rule, onUpdate, onReprompt }: RuleCardProps) {
	const [isEditing, setIsEditing] = useState(false);
	const [isReprompting, setIsReprompting] = useState(false);
	const [editedDescription, setEditedDescription] = useState(rule.description);
	const [repromptFeedback, setRepromptFeedback] = useState("");
	const [isLoading, setIsLoading] = useState(false);

	const handleSaveEdit = async () => {
		setIsLoading(true);
		try {
			await onUpdate({
				...rule,
				description: editedDescription,
				user_modified: true,
			});
			setIsEditing(false);
		} finally {
			setIsLoading(false);
		}
	};

	const handleReprompt = async () => {
		if (!repromptFeedback.trim()) return;
		setIsLoading(true);
		try {
			await onReprompt(repromptFeedback);
			setRepromptFeedback("");
			setIsReprompting(false);
		} finally {
			setIsLoading(false);
		}
	};

	const getRuleTypeColor = (type: string) => {
		const colors: Record<string, string> = {
			format: "bg-blue-100 text-blue-800",
			range: "bg-green-100 text-green-800",
			enum: "bg-purple-100 text-purple-800",
			pattern: "bg-orange-100 text-orange-800",
			sequence: "bg-pink-100 text-pink-800",
			text: "bg-gray-100 text-gray-800",
		};
		return colors[type] || colors.text;
	};

	return (
		<div className="border rounded-lg p-4 bg-white shadow-sm">
			<div className="flex items-start justify-between gap-4">
				<div className="flex-1">
					<div className="flex items-center gap-2 mb-2">
						<h4 className="font-medium text-sm">{rule.column_name}</h4>
						<span
							className={`px-2 py-0.5 rounded-full text-xs font-medium ${getRuleTypeColor(rule.rule_type)}`}
						>
							{rule.rule_type}
						</span>
						{rule.user_modified && (
							<span className="px-2 py-0.5 rounded-full text-xs font-medium bg-yellow-100 text-yellow-800">
								Modified
							</span>
						)}
					</div>

					{isEditing ? (
						<div className="space-y-2">
							<Textarea
								value={editedDescription}
								onChange={(e) => setEditedDescription(e.target.value)}
								className="text-sm"
								rows={2}
							/>
							<div className="flex gap-2">
								<Button size="sm" onClick={handleSaveEdit} disabled={isLoading}>
									{isLoading ? "Saving..." : "Save"}
								</Button>
								<Button
									size="sm"
									variant="outline"
									onClick={() => {
										setIsEditing(false);
										setEditedDescription(rule.description);
									}}
								>
									Cancel
								</Button>
							</div>
						</div>
					) : (
						<p className="text-sm text-gray-600">{rule.description}</p>
					)}

					{rule.examples && rule.examples.length > 0 && !isEditing && (
						<div className="mt-2">
							<span className="text-xs text-gray-500">Examples: </span>
							<span className="text-xs text-gray-700">
								{rule.examples.slice(0, 3).join(", ")}
							</span>
						</div>
					)}

					{rule.confidence && !isEditing && (
						<div className="mt-1 flex items-center gap-1">
							<span className="text-xs text-gray-500">Confidence:</span>
							<div className="w-20 h-1.5 bg-gray-200 rounded-full overflow-hidden">
								<div
									className="h-full bg-green-500 rounded-full"
									style={{ width: `${rule.confidence * 100}%` }}
								/>
							</div>
							<span className="text-xs text-gray-500">
								{Math.round(rule.confidence * 100)}%
							</span>
						</div>
					)}
				</div>

				{!isEditing && !isReprompting && (
					<div className="flex gap-1">
						<Button
							size="sm"
							variant="ghost"
							onClick={() => setIsEditing(true)}
							className="text-xs"
						>
							Edit
						</Button>
						<Button
							size="sm"
							variant="ghost"
							onClick={() => setIsReprompting(true)}
							className="text-xs"
						>
							Refine
						</Button>
					</div>
				)}
			</div>

			{isReprompting && (
				<div className="mt-3 pt-3 border-t space-y-2">
					<Textarea
						placeholder="Describe how to change this rule (e.g., 'Use international phone format' or 'Add more variety to names')"
						value={repromptFeedback}
						onChange={(e) => setRepromptFeedback(e.target.value)}
						className="text-sm"
						rows={2}
					/>
					<div className="flex gap-2">
						<Button
							size="sm"
							onClick={handleReprompt}
							disabled={isLoading || !repromptFeedback.trim()}
						>
							{isLoading ? "Refining..." : "Apply Changes"}
						</Button>
						<Button
							size="sm"
							variant="outline"
							onClick={() => {
								setIsReprompting(false);
								setRepromptFeedback("");
							}}
						>
							Cancel
						</Button>
					</div>
				</div>
			)}
		</div>
	);
}
