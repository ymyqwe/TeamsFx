import { exit, argv, env } from "process";
import { readdir, ensureDir, writeFile, readFile } from "fs-extra";
import { join } from "path";
import { render } from "mustache";
import { getOctokit } from "@actions/github";

import dotenv from "dotenv";
import puppeteer from "puppeteer";

dotenv.config();

const TemplateSuffix = ".mustache";
const TemplateFolder = "templates";
const TargetFolder = "mustache-output";

function toDurationMins(from: string, to: string): string {
  const fromDate = new Date(from);
  const toDate = new Date(to);
  return (
    new Date(Math.abs(toDate.getTime() - fromDate.getTime()))
      .getMinutes()
      .toString() + " mins"
  );
}

interface TemplateData {
  total_passed: number;
  total_failed: number;
  total_duration: string;
  total_cases: number;
  calendar: string;
  cases: {
    file: string;
    passed: boolean;
    author: string;
    duration: string;
    testPlanCaseId?: number;
  }[];
}

/**
 * @param github action id
 */
async function gen() {
  if (argv.length != 3) {
    throw new Error("Please input the action id");
  }

  if (!env.GITHUB_TOKEN) {
    throw new Error("Please set GITHUB_TOKEN");
  }

  let templateData: Partial<TemplateData> = {};

  const client = getOctokit(env.GITHUB_TOKEN);
  const run = await client.rest.actions.getWorkflowRun({
    owner: "OfficeDev",
    repo: "TeamsFx",
    run_id: parseInt(argv[2]),
  });

  templateData.total_duration = toDurationMins(
    run.data.created_at,
    run.data.updated_at
  );

  const jobs = await client.rest.actions.listJobsForWorkflowRun({
    owner: "OfficeDev",
    repo: "TeamsFx",
    run_id: parseInt(argv[2]),
    filter: "all",
    per_page: 100,
  });

  templateData.total_cases = jobs.data.total_count - 2;
  templateData.calendar = "12PM Beijing Time Everyday";
  templateData.total_passed = jobs.data.jobs.filter((p) => {
    return p.conclusion == "success";
  }).length;
  templateData.total_failed = jobs.data.jobs.filter((p) => {
    return p.conclusion == "failure";
  }).length;

  templateData.cases = jobs.data.jobs.map((job) => {
    return {
      file: job.name,
      passed: job.conclusion == "success",
      duration: toDurationMins(job.started_at, job.completed_at!),
      author: "",
      testPlanCaseId: 123,
    };
  });

  console.log(templateData);

  const browser = await puppeteer.launch();

  await ensureDir(join(__dirname, TargetFolder));

  const templates = await readdir(join(__dirname, TemplateFolder));
  for (const t of templates) {
    const name = t.replace(TemplateSuffix, "");
    const path = join(__dirname, TemplateFolder, t);
    const rawTemplate = (await readFile(path)).toString();
    const rendered = render(rawTemplate, templateData);
    const htmlPath = join(__dirname, TargetFolder, `${name}.html`);
    await writeFile(htmlPath, rendered);

    const page = await browser.newPage();
    await page.setViewport({
      width: 540,
      height: 400,
    });
    await page.goto(`file:${htmlPath}`);
    await page.waitForTimeout(5 * 1000);
    await page.screenshot({
      type: "png",
      path: join(__dirname, TargetFolder, `${name}.png`),
      omitBackground: true,
      //fullPage: true,
    });
  }

  await browser.close();
}

gen()
  .then(() => {
    console.log("Generation done!");
  })
  .catch((e) => {
    console.error(e);
    exit(1);
  });
