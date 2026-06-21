window.addEventListener("load", drawConnections);
window.addEventListener("resize", drawConnections);

function drawConnections() {

    const svg = document.querySelector(".connections");
    const architecture = document.querySelector(".architecture");
    const cartao = document.querySelector(".cartao-box");

    svg.innerHTML = "";

    const archRect = architecture.getBoundingClientRect();
    const cartaoRect = cartao.getBoundingClientRect();

    // Pontos do Cartão Dasa+

    const centerTop = {
        x: cartaoRect.left + cartaoRect.width / 2 - archRect.left,
        y: cartaoRect.top - archRect.top
    };

    const centerBottom = {
        x: cartaoRect.left + cartaoRect.width / 2 - archRect.left,
        y: cartaoRect.bottom - archRect.top
    };

    const topLeft = {
        x: cartaoRect.left - archRect.left,
        y: cartaoRect.top - archRect.top
    };

    const topRight = {
        x: cartaoRect.right - archRect.left,
        y: cartaoRect.top - archRect.top
    };

    const bottomLeft = {
        x: cartaoRect.left - archRect.left,
        y: cartaoRect.bottom - archRect.top
    };

    const bottomRight = {
        x: cartaoRect.right - archRect.left,
        y: cartaoRect.bottom - archRect.top
    };

    createConnection(
        centerTop,
        document.querySelector(".nav-box")
    );

    createConnection(
        topLeft,
        document.querySelector(".one-data")
    );

    createConnection(
        topRight,
        document.querySelector(".deep-water")
    );

    createConnection(
        bottomLeft,
        document.querySelector(".mdm")
    );

    createConnection(
        bottomRight,
        document.querySelector(".pagamentos")
    );

    createConnection(
        centerBottom,
        document.querySelector(".agenda")
    );

    createConnection(
        centerBottom,
        document.querySelector(".provider")
    );


    function createConnection(start, element) {

        const rect = element.getBoundingClientRect();

        const end = {
            x: rect.left + rect.width / 2 - archRect.left,
            y: rect.top + rect.height / 2 - archRect.top
        };

        const controlX = (start.x + end.x) / 2;
        const controlY = (start.y + end.y) / 2;

        const path = document.createElementNS(
            "http://www.w3.org/2000/svg",
            "path"
        );

        path.setAttribute(
            "d",
            `M ${start.x} ${start.y}
             Q ${controlX} ${controlY}
             ${end.x} ${end.y}`
        );

        path.setAttribute("class", "flow-line");

        svg.appendChild(path);
    }
}
