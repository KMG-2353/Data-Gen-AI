import { useEffect, useState } from "react";

function App() {
	const [message, setMessage] = useState("");

	useEffect(() => {
		fetch("/api/hello")
			.then((response) => response.json())
			.then((data) => setMessage(data.message));
	}, []);

	return (
		<div className="flex h-screen w-screen justify-center items-center">
			<h1 className="text-3xl font-bold">{message}</h1>
		</div>
	);
}

export default App;
