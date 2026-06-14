import argparse
import collections
import json
from pathlib import Path


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("log", type=Path)
    args = parser.parse_args()

    counts = collections.Counter()
    latencies = []

    with args.log.open(encoding="utf-8") as handle:
        for line in handle:
            if not line.strip():
                continue
            row = json.loads(line)
            counts[row.get("event", "unknown")] += 1
            if row.get("latencyMs") is not None:
                latencies.append(int(row["latencyMs"]))

    print("events")
    for event, count in counts.most_common():
        print(f"{event}: {count}")

    if latencies:
        print(f"latency_ms_avg: {sum(latencies) / len(latencies):.1f}")
        print(f"latency_ms_max: {max(latencies)}")


if __name__ == "__main__":
    main()
