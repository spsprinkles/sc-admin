import { Configuration } from "./cfg";
import { register } from "./jslink";
import { Menu } from "./menu";
import "./styles.css";

// Global Variable
window["SCAdmin"] = {
    Configuration,
    Menu,
    // TODO WebPart
    register
}

// Register the JSLink
register();