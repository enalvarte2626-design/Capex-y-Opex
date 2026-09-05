/** Icono simple de "inversión/capital" (barras crecientes) para la categoría CAPEX. */
export default function IconoCapex({ size = 17 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path
        d="M3 16.5V11M8 16.5V7M13 16.5V9.5M17 16.5V4M3 8.5L8 4.5L13 6.5L17 2.5"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinecap="round"
        strokeLinejoin="round"
      />
    </svg>
  );
}
