"use client";
import dynamic from "next/dynamic";
const Page = dynamic(() => import("../PageComponent"), { ssr: false });
export default Page;
