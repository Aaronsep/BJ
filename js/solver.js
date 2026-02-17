export function solveOptimalWithFixed(jobsMinutes, teamCount, quant) {
  const q = Math.max(1, quant | 0);

  const fixedLoads = Array(teamCount).fill(0);
  const fixedAssign = Array.from({ length: teamCount }, () => []);
  const free = [];

  for (const job of jobsMinutes) {
    if (job.fixedTeam && job.fixedTeam >= 1 && job.fixedTeam <= teamCount) {
      const idx = job.fixedTeam - 1;
      fixedLoads[idx] += job.minutes;
      fixedAssign[idx].push({ name: job.name, minutes: job.minutes });
    } else {
      free.push(job);
    }
  }

  const baseW = fixedLoads.map((minutes) => Math.round(minutes / q));
  for (let i = 0; i < teamCount; i++) {
    if (fixedLoads[i] > 0 && baseW[i] === 0) baseW[i] = 1;
  }

  const freeItems = free
    .map((job) => ({
      name: job.name,
      minutes: job.minutes,
      w: Math.max(1, Math.round(job.minutes / q)),
    }))
    .sort((a, b) => b.w - a.w);

  const itemCount = freeItems.length;

  const loads0 = baseW.slice();
  const assign0 = Array.from({ length: teamCount }, () => []);
  for (const item of freeItems) {
    let best = 0;
    for (let i = 1; i < teamCount; i++) {
      if (loads0[i] < loads0[best]) best = i;
    }
    loads0[best] += item.w;
    assign0[best].push(item);
  }

  let bestMakespanW = Math.max(...loads0);
  let bestAssign = assign0.map((arr) => arr.slice());

  const sumBase = baseW.reduce((sum, x) => sum + x, 0);
  const sumFree = freeItems.reduce((sum, x) => sum + x.w, 0);
  const avgLB = Math.ceil((sumBase + sumFree) / teamCount);
  const baseMax = Math.max(...baseW);

  const loads = baseW.slice();
  const assign = Array.from({ length: teamCount }, () => []);
  const seen = new Set();

  function key(idx) {
    const sorted = loads.slice().sort((a, b) => a - b).join(",");
    return idx + "|" + sorted;
  }

  function dfs(idx) {
    if (idx === itemCount) {
      const mk = Math.max(...loads);
      if (mk < bestMakespanW) {
        bestMakespanW = mk;
        bestAssign = assign.map((arr) => arr.slice());
      }
      return;
    }

    const currentMax = Math.max(...loads);
    const lb = Math.max(avgLB, baseMax, currentMax);
    if (lb >= bestMakespanW) return;

    const k = key(idx);
    if (seen.has(k)) return;
    seen.add(k);

    const item = freeItems[idx];
    const triedLoads = new Set();
    const order = Array.from({ length: teamCount }, (_, i) => i).sort((a, b) => loads[a] - loads[b]);

    for (const i of order) {
      const previousLoad = loads[i];
      if (triedLoads.has(previousLoad)) continue;
      triedLoads.add(previousLoad);

      const newLoad = previousLoad + item.w;
      if (newLoad >= bestMakespanW) continue;

      loads[i] = newLoad;
      assign[i].push(item);
      dfs(idx + 1);
      assign[i].pop();
      loads[i] = previousLoad;
    }
  }

  dfs(0);

  const teams = Array.from({ length: teamCount }, (_, i) => ({
    jobs: [...fixedAssign[i]],
    totalMinutes: fixedLoads[i],
  }));

  for (let i = 0; i < teamCount; i++) {
    for (const item of bestAssign[i]) {
      teams[i].jobs.push({ name: item.name, minutes: item.minutes });
      teams[i].totalMinutes += item.minutes;
    }
    teams[i].jobs.sort((a, b) => b.minutes - a.minutes);
  }

  return { teams };
}
