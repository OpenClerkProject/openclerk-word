import { BluebookRuleSet } from "./types";

class BluebookRuleSetRegistry {
  private ruleSets = new Map<string, BluebookRuleSet>();

  register(ruleSet: BluebookRuleSet): void {
    this.ruleSets.set(ruleSet.id, ruleSet);
  }

  unregister(id: string): void {
    this.ruleSets.delete(id);
  }

  get(id: string): BluebookRuleSet | undefined {
    return this.ruleSets.get(id);
  }

  list(): BluebookRuleSet[] {
    return Array.from(this.ruleSets.values());
  }
}

export const bluebookRuleSetRegistry = new BluebookRuleSetRegistry();
export { BluebookRuleSetRegistry };
