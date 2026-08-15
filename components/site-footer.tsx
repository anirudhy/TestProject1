export function SiteFooter() {
  return (
    <footer className="border-t border-border py-8 text-sm text-muted-foreground">
      <div className="mx-auto max-w-6xl px-4 space-y-2">
        <p>
          PerkStack is an unofficial, community-maintained directory. It is not affiliated with or endorsed by the
          companies listed. Data may be inaccurate or outdated — verify with your employer before making decisions.
        </p>
        <p>Not tax or investment advice.</p>
        <p className="pt-2">© {new Date().getFullYear()} PerkStack</p>
      </div>
    </footer>
  );
}
