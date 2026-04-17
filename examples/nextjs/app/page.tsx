import Link from "next/link";

export default function Page() {
  return (
    <div>
      <h1>wasm-xlsxwriter Next.js Integration Test</h1>
      <ul>
        <li>
          <Link href="/app-client">App Router - Client Side</Link>
        </li>
        <li>
          <a href="/api/generate">App Router - Route Handler (Server)</a>
        </li>
        <li>
          <Link href="/pages-client">Pages Router - Client Side</Link>
        </li>
        <li>
          <a href="/api/pages-generate">Pages Router - API Route (Server)</a>
        </li>
      </ul>
    </div>
  );
}
