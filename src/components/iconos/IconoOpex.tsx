/** Icono simple de "gasto recurrente/factura" para la categoría OPEX. */
export default function IconoOpex({ size = 17 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
      <path
        d="M5 2.5H15V17.5L12.5 16L10 17.5L7.5 16L5 17.5V2.5Z"
        stroke="currentColor"
        strokeWidth="1.6"
        strokeLinejoin="round"
      />
      <path d="M7.5 6.5H12.5M7.5 9.5H12.5" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
    </svg>
  );
}
