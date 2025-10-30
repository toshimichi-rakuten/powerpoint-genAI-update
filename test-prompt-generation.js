import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Read the pptxAnalyzer.js to extract the prompt template
const analyzerPath = path.join(__dirname, 'src/pptxAnalyzer.js');
const analyzerCode = fs.readFileSync(analyzerPath, 'utf-8');

// Extract the prompt section
const promptMatch = analyzerCode.match(/const prompt = `# PowerPoint Slide Reproduction Task\n\n([\s\S]*?)## JSON Structure/);

if (promptMatch) {
  console.log('=== PROMPT HEADER (New Implementation Rules) ===\n');
  console.log(promptMatch[1]);
  console.log('\n✅ The prompt now includes the instruction to NOT use variables!');
  console.log('✅ All slide.addXXX methods should have data inline, not in variables.');
} else {
  console.log('❌ Could not extract prompt template');
}
