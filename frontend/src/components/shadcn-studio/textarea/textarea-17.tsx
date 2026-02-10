import { useId, useState } from "react";
import { Label } from "@/components/ui/label";
import { Textarea } from "@/components/ui/textarea";

type specialinstrction = {
	value: string;
	inputchange: (newvalue: string) => void;
};
const TextareaAutoGrowDemo = ({ value, inputchange }: specialinstrction) => {
	const handlechange = (event: React.ChangeEvent<HTMLTextAreaElement>) => {
		const newvalue = event.target.value;
		inputchange(newvalue);
	};

	const id = useId();
	return (
		<div className="w-full  space-y-2">
			<Label htmlFor={id}>Important Rules</Label>
			<Textarea
				id={id}
				placeholder="Special Instructions"
				className="field-sizing-content max-h-30 min-h-0 resize-none py-1.75"
				onChange={handlechange}
			/>
		</div>
	);
};

export default TextareaAutoGrowDemo;
