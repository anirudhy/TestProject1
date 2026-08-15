import Link from "next/link";

const links = [
  { href: "/companies", label: "Companies" },
  { href: "/compare", label: "Compare" },
  { href: "/calculator", label: "Calculator" },
  { href: "/contribute", label: "Contribute" },
];

export function SiteHeader() {
  return (
    <header className="border-b border-border bg-background/95 backdrop-blur supports-[backdrop-filter]:bg-background/60 sticky top-0 z-40">
      <div className="mx-auto flex h-14 max-w-6xl items-center justify-between px-4">
        <Link href="/" className="font-semibold tracking-tight">
          Perk<span className="text-primary">Stack</span>
        </Link>
        <nav className="flex items-center gap-6 text-sm">
          {links.map((link) => (
            <Link key={link.href} href={link.href} className="text-muted-foreground transition-colors hover:text-foreground">
              {link.label}
            </Link>
          ))}
        </nav>
      </div>
    </header>
  );
}
