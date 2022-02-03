import { ContextInfo } from "gd-sprest-bs";
import { Configuration } from "./cfg";
import { register } from "./jslink";
import "./styles.css";

// Global Variable
window["SCAdmin"] = { Configuration, register }

// See if this is a classic page and a list is detected
if (ContextInfo.existsFl && ContextInfo.listId) {
    // Register the jslink
    register();
}