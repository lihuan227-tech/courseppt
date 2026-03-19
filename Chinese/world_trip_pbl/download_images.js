/**
 * Download images for Day 1: Asia (China, Japan, India)
 * Uses https with manual redirect following and User-Agent header.
 * Run: node download_images.js
 */
const https = require("https");
const http = require("http");
const fs = require("fs");
const path = require("path");

const IMAGES_DIR = path.join(__dirname, "images");
if (!fs.existsSync(IMAGES_DIR)) fs.mkdirSync(IMAGES_DIR, { recursive: true });

const UA = "Mozilla/5.0 (educational project)";

const DOWNLOADS = [
  // China
  {
    file: "china_great_wall.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/10/20090529_Great_Wall_8185.jpg/1280px-20090529_Great_Wall_8185.jpg"],
  },
  {
    file: "china_forbidden_city.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/c9/Beijing-Forbidden_City1.jpg/1280px-Beijing-Forbidden_City1.jpg"],
  },
  {
    file: "china_dumplings.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/4/4c/Buuz.jpg/1280px-Buuz.jpg"],
  },
  {
    file: "china_beijing.jpg",
    urls: [
      "https://upload.wikimedia.org/wikipedia/commons/thumb/2/22/Beijing_CBD_2024.jpg/1280px-Beijing_CBD_2024.jpg",
      "https://upload.wikimedia.org/wikipedia/commons/thumb/f/f5/Beijing_montage_1.png/800px-Beijing_montage_1.png",
    ],
  },
  {
    file: "china_spring_festival.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/c/c4/2012_Chinese_New_Year_dragon.jpg/1280px-2012_Chinese_New_Year_dragon.jpg"],
  },
  {
    file: "china_inventions.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/d/d4/Compass_in_a_wooden_box.jpg/1024px-Compass_in_a_wooden_box.jpg"],
  },
  // Japan
  {
    file: "japan_fuji.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/1/1b/080103_hakridge_fuji.jpg/1280px-080103_hakridge_fuji.jpg"],
  },
  {
    file: "japan_sushi.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/6/60/Sushi_platter.jpg/1280px-Sushi_platter.jpg"],
  },
  {
    file: "japan_tokyo.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/b2/Skyscrapers_of_Shinjuku_2009_January.jpg/1280px-Skyscrapers_of_Shinjuku_2009_January.jpg"],
  },
  {
    file: "japan_cherry_blossom.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/7/7a/Yoshino_cherry_blossoms_and_the_National_Diet_Building.jpg/1280px-Yoshino_cherry_blossoms_and_the_National_Diet_Building.jpg"],
  },
  {
    file: "japan_pokemon.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/b7/Pok%C3%A9mon_Center_Mega_Tokyo.jpg/1280px-Pok%C3%A9mon_Center_Mega_Tokyo.jpg"],
  },
  // India
  {
    file: "india_taj_mahal.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/b/bd/Taj_Mahal%2C_Agra%2C_India_edit3.jpg/1280px-Taj_Mahal%2C_Agra%2C_India_edit3.jpg"],
  },
  {
    file: "india_holi.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/3/3e/Holi_celebrations%2C_bystanders_drenched_in_colors.jpg/1280px-Holi_celebrations%2C_bystanders_drenched_in_colors.jpg"],
  },
  {
    file: "india_ganges.jpg",
    urls: ["https://upload.wikimedia.org/wikipedia/commons/thumb/a/a0/Varanasi_-_Ganges_River_-_002.jpg/1280px-Varanasi_-_Ganges_River_-_002.jpg"],
  },
];

function download(url, destPath, maxRedirects = 5) {
  return new Promise((resolve, reject) => {
    if (maxRedirects <= 0) return reject(new Error("Too many redirects"));

    const mod = url.startsWith("https") ? https : http;
    const req = mod.get(url, { headers: { "User-Agent": UA } }, (res) => {
      if ([301, 302, 307].includes(res.statusCode)) {
        const loc = res.headers.location;
        if (!loc) return reject(new Error("Redirect with no location"));
        const next = loc.startsWith("http") ? loc : new URL(loc, url).href;
        res.resume();
        return download(next, destPath, maxRedirects - 1).then(resolve, reject);
      }
      if (res.statusCode !== 200) {
        res.resume();
        return reject(new Error("HTTP " + res.statusCode));
      }
      const ws = fs.createWriteStream(destPath);
      res.pipe(ws);
      ws.on("finish", () => ws.close(() => resolve()));
      ws.on("error", reject);
    });
    req.on("error", reject);
    req.setTimeout(30000, () => { req.destroy(); reject(new Error("Timeout")); });
  });
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

async function main() {
  const results = [];
  for (const item of DOWNLOADS) {
    const dest = path.join(IMAGES_DIR, item.file);
    let ok = false;
    for (const url of item.urls) {
      try {
        console.log("Downloading " + item.file + " ...");
        await download(url, dest);
        const stat = fs.statSync(dest);
        if (stat.size < 1000) {
          fs.unlinkSync(dest);
          throw new Error("File too small (" + stat.size + " bytes)");
        }
        console.log("  OK (" + Math.round(stat.size / 1024) + " KB)");
        ok = true;
        break;
      } catch (e) {
        console.log("  FAIL: " + e.message + " — trying next URL...");
      }
    }
    results.push({ file: item.file, ok });
    if (!ok) console.log("  FAILED all URLs for " + item.file);
    await sleep(1000);
  }

  console.log("\n=== SUMMARY ===");
  const succeeded = results.filter((r) => r.ok);
  const failed = results.filter((r) => !r.ok);
  console.log("Succeeded: " + succeeded.length + "/" + results.length);
  succeeded.forEach((r) => console.log("  OK  " + r.file));
  if (failed.length) {
    console.log("Failed: " + failed.length);
    failed.forEach((r) => console.log("  FAIL  " + r.file));
  }
}

main().catch(console.error);
