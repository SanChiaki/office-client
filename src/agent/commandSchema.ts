import { z } from "zod";

export const actionSchema = z.object({
  type: z.string(),
  args: z.record(z.any()),
});

export const commandEnvelopeSchema = z.object({
  assistant_message: z.string(),
  mode: z.enum(["chat", "excel_action", "skill"]),
  skill_name: z.string().optional(),
  requires_confirmation: z.boolean().default(false),
  actions: z.array(actionSchema).default([]),
});

export type CommandEnvelope = z.infer<typeof commandEnvelopeSchema>;
