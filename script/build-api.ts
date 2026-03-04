import { build as esbuild } from "esbuild";
import { rm, readFile } from "fs/promises";

async function buildApi() {
  await rm("dist-api", { recursive: true, force: true });

  const pkg = JSON.parse(await readFile("package.json", "utf-8"));
  const allDeps = [
    ...Object.keys(pkg.dependencies || {}),
    ...Object.keys(pkg.devDependencies || {}),
  ];
  const externals = allDeps;

  await esbuild({
    entryPoints: ["server/index.ts"],
    platform: "node",
    bundle: true,
    format: "cjs",
    outfile: "dist-api/index.cjs",
    define: {
      "process.env.NODE_ENV": '"production"',
    },
    minify: true,
    external: externals,
    logLevel: "info",
  });
}

buildApi().catch((err) => {
  console.error(err);
  process.exit(1);
});
