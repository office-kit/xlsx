{
  description = "openxml-js — TypeScript port of openpyxl. Reproducible dev environment.";

  inputs = {
    nixpkgs.url = "github:NixOS/nixpkgs/nixos-unstable";
    flake-utils.url = "github:numtide/flake-utils";
  };

  outputs = { self, nixpkgs, flake-utils }:
    flake-utils.lib.eachDefaultSystem (system:
      let
        pkgs = import nixpkgs { inherit system; };
        # Pin to the same Node major used by CI / engines.node in package.json.
        nodejs = pkgs.nodejs_22;
        # Pin the pnpm major used by CI / package.json#packageManager.
        pnpm = pkgs.pnpm_11;
      in
      {
        devShells.default = pkgs.mkShell {
          name = "openxml-js-dev";
          packages = [
            nodejs
            pnpm
            pkgs.git
            # `python3` is needed by `pnpm install` for any node-gyp builds and
            # for the openpyxl reference fixtures.
            pkgs.python3
          ];
          shellHook = ''
            echo "openxml-js dev shell"
            echo "  node $(node --version)"
            echo "  pnpm $(pnpm --version)"
            echo "  python $(python3 --version)"
            export NODE_OPTIONS="''${NODE_OPTIONS:-}"
            # Use the in-repo node_modules instead of any global pnpm store
            # paths; reproducibility wins over speed for CI parity.
            export PNPM_HOME="$PWD/.pnpm"
            export PATH="$PNPM_HOME:$PATH"
          '';
        };

        # Convenience for `nix flake check` — does a fast typecheck + test run.
        # Heavier checks (build, lint) are run in CI directly.
        checks.default = pkgs.runCommand "openxml-js-check"
          { buildInputs = [ nodejs pnpm ]; }
          ''
            cp -r ${self} src
            cd src
            chmod -R u+w .
            pnpm install --frozen-lockfile --offline 2>/dev/null || pnpm install --frozen-lockfile
            pnpm typecheck
            pnpm test
            mkdir -p $out
            touch $out/ok
          '';
      });
}
