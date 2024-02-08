// Returns the logo as an SVG element
export function getLogo(height?: number, width?: number, className?: string) {
    // Get the icon element
    let elDiv = document.createElement("div");
    elDiv.innerHTML = "<svg width='32' height='32' fill='currentColor' viewBox='0 0 65 65' xmlns='http://www.w3.org/2000/svg'><g transform='matrix(0.99763673,0,0,1,-0.90450782,-1)'><path d='M 64.8,32.1 C 61.9,29.9 49.3,36.7 33.5,36.7 17.6,36.7 5.2,29.9 2.2,32.1 -2,35.3 10.8,45.4 33.5,45.4 56.8,45.4 68.8,35.1 64.8,32.1 Z'/><path d='M 54.8,31.3 V 8.9 c 0,-1.8 -1,-3.4 -2.6,-4.2 -0.6,-0.3 -1.3,-0.5 -2,-0.5 -1,0 -2,0.3 -2.9,1 L 38,12.6 c -2.6,2.1 -6.3,2.1 -8.9,0 L 19.8,5.3 C 18.4,4.2 16.5,4 14.9,4.8 13.3,5.6 12.3,7.2 12.3,9 v 22.4 c 5.5,1.5 13,3.4 21.3,3.4 8.1,-0.1 15.6,-2.1 21.2,-3.5 z'/><path d='M 60.1,50.9 H 6.9 c -0.6,0 -1,0.5 -1,1 0,0.6 0.4,1 1,1 h 2.3 c 0.5,5.5 5.1,9.8 10.8,9.8 5.6,0 10.3,-4.3 10.8,-9.8 h 5.5 c 0.5,5.5 5.1,9.8 10.8,9.8 5.6,0 10.3,-4.3 10.8,-9.8 h 2.4 c 0.6,0 1,-0.4 1,-1 -0.2,-0.5 -0.6,-1 -1.2,-1 z'/></g></svg>";
    let icon = elDiv.firstChild as SVGImageElement;
    if (icon) {
        // See if a class name exists
        if (className) {
            // Parse the class names
            let classNames = className.split(' ');
            for (let i = 0; i < classNames.length; i++) {
                // Add the class name
                icon.classList.add(classNames[i]);
            }
        }
        // Set the height/width
        height ? icon.setAttribute("height", (height).toString()) : null;
        width ? icon.setAttribute("width", (width).toString()) : null;
        // Hide the icon as non-interactive content from the accessibility API
        icon.setAttribute("aria-hidden", "true");
        // Update the styling
        icon.style.pointerEvents = "none";
        // Support for IE
        icon.setAttribute("focusable", "false");
    }
    // Return the icon
    return icon;
}