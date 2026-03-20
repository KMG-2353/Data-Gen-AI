// Rule types for pattern-based data generation

export type RuleType =
	| "format"
	| "range"
	| "enum"
	| "pattern"
	| "sequence"
	| "text";

export interface ColumnRule {
	column_name: string;
	rule_type: RuleType;
	description: string;
	pattern?: string | null;
	constraints: Record<string, unknown>;
	examples: string[];
	confidence: number;
	llm_reasoning: string;
	user_modified?: boolean;
}

export interface SheetRuleSet {
	rules: ColumnRule[];
	cross_column_rules: string[];
}

export interface SheetData {
	sheet_name: string;
	headers: string[];
	unique_headers: string[];
	sample_count: number;
	total_rows: number;
}

export interface UploadResponse {
	session_id: string;
	sheets: SheetData[];
	sheet_names: string[];
	sheet_count: number;
	filename: string;
}

export interface AnalysisProgress {
	current_sheet: string | null;
	sheets_completed: string[];
	sheets_total: number;
	progress: number;
	status: "idle" | "analyzing" | "complete" | "error";
	error_message?: string;
}

export interface AnalysisEvent {
	event: "sheet_start" | "sheet_complete" | "error" | "complete";
	sheet_name?: string;
	progress: number;
	message?: string;
	rules?: SheetRuleSet;
}

export interface RulesResponse {
	session_id: string;
	rule_sets: Record<string, SheetRuleSet>;
	sheets: string[];
}

export interface UpdateRuleResponse {
	success: boolean;
	sheet_name: string;
	column_name: string;
	updated_rule: ColumnRule;
}

export interface RepromptResponse {
	session_id: string;
	sheet_name: string;
	column_name: string;
	old_rule: ColumnRule;
	new_rule: ColumnRule;
}

export interface GenerateResponse {
	session_id: string;
	sheets_generated: string[];
	row_count_per_sheet: number;
	status: string;
}
