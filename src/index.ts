import { Configuration } from "./cfg";
import { register } from "./jslink";
import "./styles.css";

// Global Variable
window["SCAdmin"] = { Configuration, register }

// Register the jslink
register();