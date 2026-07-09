import React from "react";
import ReactDOM from "react-dom/client";
import LemburApp from "./LemburApp.jsx";
import "./index.css";

const savedTheme = localStorage.getItem("lembur:theme") || "light";
if (savedTheme === "dark") document.documentElement.classList.add("dark");
else document.documentElement.classList.remove("dark");

ReactDOM.createRoot(document.getElementById("root")).render(<LemburApp />);
