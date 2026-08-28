from __future__ import annotations

import argparse
import subprocess
import tarfile
import zipfile
from email import message_from_bytes
from email.message import Message
from pathlib import Path

import tomllib
from packaging.utils import parse_sdist_filename, parse_wheel_filename
from packaging.version import Version


def project_version(root: Path) -> str:
    with (root / "pyproject.toml").open("rb") as stream:
        project = tomllib.load(stream)["project"]
    version = str(project["version"])
    if str(Version(version)) != version:
        raise ValueError(f"Project version is not canonical PEP 440: {version!r}")
    return version


def verify_release_checkout(root: Path, *, tag: str, expected_commit: str) -> None:
    version = project_version(root)
    if tag != version:
        raise ValueError(f"Release tag {tag!r} does not match project {version!r}")
    head = _git(root, "rev-parse", "HEAD")
    tag_commit = _git(root, "rev-parse", f"refs/tags/{tag}^{{commit}}")
    if head != tag_commit or expected_commit != tag_commit:
        raise ValueError(
            "Release checkout, event commit, and tag commit must be identical: "
            f"head={head}, event={expected_commit}, tag={tag_commit}"
        )
    main_ancestor = subprocess.run(
        ["git", "merge-base", "--is-ancestor", tag_commit, "origin/main"],
        cwd=root,
        check=False,
    )
    if main_ancestor.returncode != 0:
        raise ValueError(
            f"Release commit is not contained in origin/main: {tag_commit}"
        )
    if _git(root, "status", "--porcelain=v1", "--untracked-files=all"):
        raise ValueError("Release checkout must be clean")


def verify_distributions(root: Path, dist: Path) -> tuple[Path, Path]:
    version = project_version(root)
    entries = sorted(dist.iterdir()) if dist.is_dir() else []
    if any(not entry.is_file() or entry.is_symlink() for entry in entries):
        raise ValueError("Distribution directory must contain regular files only")
    wheels = [entry for entry in entries if entry.suffix == ".whl"]
    sdists = [entry for entry in entries if entry.name.endswith(".tar.gz")]
    if len(entries) != 2 or len(wheels) != 1 or len(sdists) != 1:
        raise ValueError("Distribution must contain exactly one wheel and one sdist")

    identities = {
        wheels[0]: _wheel_identity(wheels[0]),
        sdists[0]: _sdist_identity(sdists[0]),
    }
    for artifact, (name, artifact_version) in identities.items():
        if name != "docxrender" or artifact_version != version:
            raise ValueError(
                f"Unexpected distribution identity for {artifact.name}: "
                f"{name} {artifact_version}; expected docxrender {version}"
            )
    wheel_name, wheel_version, _, _ = parse_wheel_filename(wheels[0].name)
    sdist_name, sdist_version = parse_sdist_filename(sdists[0].name)
    if wheel_name != "docxrender" or str(wheel_version) != version:
        raise ValueError(f"Wheel filename does not match project identity: {wheels[0]}")
    if sdist_name != "docxrender" or str(sdist_version) != version:
        raise ValueError(f"Sdist filename does not match project identity: {sdists[0]}")
    return wheels[0], sdists[0]


def _wheel_identity(path: Path) -> tuple[str, str]:
    with zipfile.ZipFile(path) as archive:
        names = [
            name for name in archive.namelist() if name.endswith(".dist-info/METADATA")
        ]
        if len(names) != 1:
            raise ValueError(f"Wheel must contain one METADATA file: {path}")
        metadata = message_from_bytes(archive.read(names[0]))
    return _identity(metadata, path)


def _sdist_identity(path: Path) -> tuple[str, str]:
    with tarfile.open(path, mode="r:gz") as archive:
        members = [
            member
            for member in archive.getmembers()
            if member.isfile() and member.name.endswith("/PKG-INFO")
        ]
        if len(members) != 1:
            raise ValueError(f"Sdist must contain one PKG-INFO file: {path}")
        stream = archive.extractfile(members[0])
        if stream is None:
            raise ValueError(f"Cannot read sdist metadata: {path}")
        metadata = message_from_bytes(stream.read())
    return _identity(metadata, path)


def _identity(metadata: Message, path: Path) -> tuple[str, str]:
    name = metadata.get("Name")
    version = metadata.get("Version")
    if name is None or version is None:
        raise ValueError(f"Distribution metadata lacks Name or Version: {path}")
    return str(name), str(version)


def _git(root: Path, *arguments: str) -> str:
    return subprocess.run(
        ["git", *arguments],
        cwd=root,
        check=True,
        capture_output=True,
        text=True,
    ).stdout.strip()


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--dist", type=Path, required=True)
    parser.add_argument("--tag")
    parser.add_argument("--expected-commit")
    arguments = parser.parse_args()
    if (arguments.tag is None) != (arguments.expected_commit is None):
        parser.error("--tag and --expected-commit must be supplied together")
    root = Path.cwd()
    if arguments.tag is not None and arguments.expected_commit is not None:
        verify_release_checkout(
            root,
            tag=arguments.tag,
            expected_commit=arguments.expected_commit,
        )
    wheel, sdist = verify_distributions(root, arguments.dist)
    print(f"verified release artifacts: {wheel.name}, {sdist.name}")


if __name__ == "__main__":
    main()
