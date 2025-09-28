import React from "react";
import ReactDOM from "react-dom/client";
import "./index.css"; // optional, if you want Tailwind or your own styles
import App from "./App";

const rootElement = document.getElementById("root");
if (!rootElement) throw new Error("Failed to find the root element");

const root = ReactDOM.createRoot(rootElement);
root.render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
