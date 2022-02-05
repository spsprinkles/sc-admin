import { Configuration } from "./cfg";
import { register } from "./jslink";
import { Dashboard } from "./dashboard";
import { Menu } from "./menu";
import "./styles.css";

// Global Variable
window["SCAdmin"] = {
    Configuration,
    Dashboard,
    Menu,
    register,
    render: (el, context) => { new Dashboard(el, context); }
}

// Register the JSLink
register();