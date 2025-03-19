tabs.forEach(tab => {
    tab.addEventListener("click", function () {
        submenus.forEach(menu => menu.classList.remove("mostrar"));

        const selectedTab = this.getAttribute("data-tab");
        if (selectedTab === "datos") {
            submenuDatos.classList.add("mostrar");
        }
    });
});
