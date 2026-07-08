import { bluebookRuleSetRegistry } from "./registry";
import { Bluebook20thEdition } from "./edition20th";
import { Bluebook21stEdition } from "./edition21st";
import { Bluebook22ndEdition } from "./edition22nd";

bluebookRuleSetRegistry.register(new Bluebook22ndEdition());
bluebookRuleSetRegistry.register(new Bluebook21stEdition());
bluebookRuleSetRegistry.register(new Bluebook20thEdition());

export { bluebookRuleSetRegistry } from "./registry";
export * from "./types";
