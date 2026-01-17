// Theme Toggle Script
(function() {
    const html = document.documentElement;

    // Check for saved theme preference or respect system preference
    function getPreferredTheme() {
        const saved = localStorage.getItem('theme');
        if (saved) return saved;
        return globalThis.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
    }

    function setTheme(theme) {
        html.dataset.theme = theme;
        localStorage.setItem('theme', theme);

        const themeToggle = document.getElementById('themeToggle');
        if (themeToggle) {
            themeToggle.textContent = theme === 'dark' ? '\u2600\uFE0F' : '\uD83C\uDF19';
            themeToggle.title = theme === 'dark' ? '\u5207\u63DB\u81F3\u4EAE\u8272\u4E3B\u984C' : '\u5207\u63DB\u81F3\u6697\u8272\u4E3B\u984C';
        }
    }

    // Apply theme immediately to avoid flash
    setTheme(getPreferredTheme());

    // Wait for DOM to be ready for button binding
    document.addEventListener('DOMContentLoaded', function() {
        // Theme toggle binding
        const themeToggle = document.getElementById('themeToggle');
        if (themeToggle) {
            // Update button state
            const currentTheme = html.dataset.theme || 'light';
            themeToggle.textContent = currentTheme === 'dark' ? '\u2600\uFE0F' : '\uD83C\uDF19';
            themeToggle.title = currentTheme === 'dark' ? '\u5207\u63DB\u81F3\u4EAE\u8272\u4E3B\u984C' : '\u5207\u63DB\u81F3\u6697\u8272\u4E3B\u984C';

            // Toggle theme on button click
            themeToggle.addEventListener('click', function() {
                const current = html.dataset.theme;
                setTheme(current === 'dark' ? 'light' : 'dark');
            });
        }

        // Mobile menu toggle
        const menuToggle = document.querySelector('.menu-toggle');
        const navMenu = document.getElementById('navMenu');

        if (menuToggle && navMenu) {
            menuToggle.addEventListener('click', function() {
                navMenu.classList.toggle('active');
                menuToggle.textContent = navMenu.classList.contains('active') ? '\u2715' : '\u2630';
            });

            // Close menu when clicking menu links (mobile)
            const menuLinks = navMenu.querySelectorAll('a');
            menuLinks.forEach(function(link) {
                link.addEventListener('click', function() {
                    if (globalThis.innerWidth <= 768) {
                        navMenu.classList.remove('active');
                        menuToggle.textContent = '\u2630';
                    }
                });
            });
        }
    });

    // Listen for system theme changes
    globalThis.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', function(e) {
        if (!localStorage.getItem('theme')) {
            setTheme(e.matches ? 'dark' : 'light');
        }
    });
})();
