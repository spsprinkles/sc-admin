const icons = new Map<string, string>([
    ["EntryView", "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='32' fill='none' viewBox='0 0 2048 2048'><path d='M777 1920q15 35 36 67t48 61H0V0h1536v911q-32-9-64-17t-64-13V128H128v1792h649zm503-1408H256V384h1024v128zM256 768h1024v128H256V768zm960 640q66 0 124 25t101 69 69 102 26 124q0 66-25 124t-69 102-102 69-124 25q-66 0-124-25t-102-68-69-102-25-125q0-66 25-124t68-101 102-69 125-26zm0 512q40 0 75-15t61-41 41-61 15-75q0-40-15-75t-41-61-61-41-75-15q-40 0-75 15t-61 41-41 61-15 75q0 40 15 75t41 61 61 41 75 15zm0-896q100 0 200 21t193 64 173 103 139 139 93 173 34 204h-128q0-91-29-169t-81-142-119-113-147-83-162-51-166-18q-82 0-166 17t-162 51-146 83-120 114-80 142-30 169H384q0-109 34-204t93-173 139-139 172-103 193-63 201-22z'></path></svg>"],
    ["ExcelDocument", "<svg xmlns='http://www.w3.org/2000/svg' width='32' height='32' fill='none' viewBox='0 0 2048 2048'><path d='M2048 475v1445q0 27-10 50t-27 40-41 28-50 10H640q-27 0-50-10t-40-27-28-41-10-50v-256H115q-24 0-44-9t-37-25-25-36-9-45V627q0-24 9-44t25-37 36-25 45-9h397V128q0-27 10-50t27-40 41-28 50-10h933q26 0 49 9t42 28l347 347q18 18 27 41t10 50zm-384-256v165h165l-165-165zM261 1424h189q2-4 12-23t25-45 29-55 29-53 23-41 10-17q27 59 60 118t65 116h187l-209-339 205-333H707q-31 57-60 114t-63 112q-29-57-57-113t-57-113H279l199 335-217 337zm379 496h1280V512h-256q-27 0-50-10t-40-27-28-41-10-50V128H640v384h397q24 0 44 9t37 25 25 36 9 45v922q0 24-9 44t-25 37-36 25-45 9H640v256zm640-1024V768h512v128h-512zm0 256v-128h512v128h-512zm0 256v-128h512v128h-512z'></path></svg>"]
]);

// Returns an icon as an SVG element
export function GetIcon(height?, width?, iconName?, className?) {
    // Get the icon element
    let elDiv = document.createElement("div");
    elDiv.innerHTML = iconName ? icons.get(iconName) : icons.get('EntryView');
    let icon = elDiv.firstChild as SVGImageElement;
    if (icon) {
        // See if a class name exists
        if (className) {
            // Add the default class name
            icon.classList.add("icon-svg");
            // Parse the class names
            let classNames = className.split(' ');
            for (let i = 0; i < classNames.length; i++) {
                // Add the class name
                icon.classList.add(classNames[i]);
            }
        } else {
            // Add the default class name
            icon.classList.add("icon-svg");
        }
        // Set the height/width
        height ? icon.setAttribute("height", (height).toString()) : null;
        width ? icon.setAttribute("width", (width).toString()) : null;
        // Update the styling
        icon.style.pointerEvents = "none";
        // Support for IE
        icon.setAttribute("focusable", "false");
    }
    // Return the icon
    return icon;
}