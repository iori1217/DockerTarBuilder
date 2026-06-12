# AGENTS.md

## Cursor Cloud specific instructions

### What this repo is
`DockerTarBuilder` is a **GitHub Actions workflow-only project** — there is no application source code, no package manager, and no database. The six workflows in `.github/workflows/` are manually triggered (`workflow_dispatch`) and run on GitHub-hosted `ubuntu-22.04` runners. Their core logic pulls Docker image(s) for a target platform and packages them as `.tar.gz` (uploaded as an Artifact, or published as a Release asset for the `*-release.yml` variants). End users load the result with `docker load -i <file>.tar.gz`.

### Lint
There is no project-specific linter. Use `actionlint` to validate the workflow YAML (pre-installed at `/usr/local/bin/actionlint` in the snapshot):

```
actionlint .github/workflows/*.yml
```

### Build / test / run
There is nothing to build and no automated test suite. The product is "run" by reproducing the workflow's core steps locally with Docker, e.g.:

```
docker pull alpine:latest --platform linux/amd64
docker save alpine:latest -o alpine_latest-amd64.tar
gzip -c alpine_latest-amd64.tar > alpine_latest-amd64.tar.gz
docker load -i alpine_latest-amd64.tar.gz   # verify the archive
```

### Docker daemon (non-obvious startup caveat)
- Docker is installed in the snapshot but the daemon is **not** auto-started (this container has no running `systemd`). Start it manually before any docker command:
  ```
  sudo dockerd        # run in a background terminal/tmux session
  ```
- The daemon is configured (`/etc/docker/daemon.json`) to use the `fuse-overlayfs` storage driver with `containerd-snapshotter` disabled — this is **required** for Docker to work in the Cloud Agent VM (kernel lacks full overlay2 support, and Docker 29's default containerd snapshotter is incompatible with fuse-overlayfs). Do not remove this config.
- `iptables` is set to `iptables-legacy` for the same VM-compatibility reason.
- Docker commands generally need `sudo` here. Note `docker save -o <file>` writes the file as `root`; run the subsequent `gzip` as root (or `chown`) or it will hit a permission error.
