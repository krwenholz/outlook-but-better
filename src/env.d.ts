// esbuild loads .css imports as text (see scripts/build.ts).
declare module "*.css" {
  const text: string;
  export default text;
}
