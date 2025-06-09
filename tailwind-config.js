tailwind.config = {
    theme: {
        extend: {
            colors: {
                primary: '#1e40af',
                'primary-dark': '#1e3a8a',
                secondary: '#3b82f6',
                accent: '#10b981',
                background: '#f8fafc',
                surface: '#ffffff',
                'surface-hover': '#f1f5f9',
                error: '#ef4444',
                warning: '#f59e0b',
                success: '#10b981',
                neutral: '#64748b',
                'neutral-light': '#94a3b8',
                'text-primary': '#0f172a',
                'text-secondary': '#475569',
                'border-light': '#e2e8f0',
                'border-medium': '#cbd5e1',
            },
            fontFamily: {
                sans: ['Inter', 'ui-sans-serif', 'system-ui', 'sans-serif'],
            },
            boxShadow: {
                'soft': '0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06)',
                'medium': '0 4px 6px -1px rgba(0, 0, 0, 0.1), 0 2px 4px -1px rgba(0, 0, 0, 0.06)',
                'large': '0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05)',
                'glass': '0 8px 32px 0 rgba(31, 38, 135, 0.1)',
            },
            animation: {
                'fade-in': 'fadeIn 0.3s ease-out',
                'slide-in': 'slideIn 0.3s ease-out',
                'bounce-subtle': 'bounceSubtle 0.6s ease-out',
                'pulse-slow': 'pulse 2s cubic-bezier(0.4, 0, 0.6, 1) infinite',
            },
            keyframes: {
                fadeIn: {
                    '0%': { opacity: '0', transform: 'translateY(10px)' },
                    '100%': { opacity: '1', transform: 'translateY(0)' },
                },
                slideIn: {
                    '0%': { transform: 'translateX(-100%)' },
                    '100%': { transform: 'translateX(0)' },
                },
                bounceSubtle: {
                    '0%, 20%, 50%, 80%, 100%': { transform: 'translateY(0)' },
                    '40%': { transform: 'translateY(-4px)' },
                    '60%': { transform: 'translateY(-2px)' },
                },
            },
        }
    }
}