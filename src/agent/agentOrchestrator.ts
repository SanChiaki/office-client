import { inferSkillRoute } from "../skills/registry";

export type RouteDecision =
  | {
      mode: "skill";
      skillName: string;
    }
  | {
      mode: "chat";
    };

export function decideRoute(input: string): RouteDecision {
  const skillRoute = inferSkillRoute(input);
  if (skillRoute) {
    return {
      mode: "skill",
      skillName: skillRoute.skillName,
    };
  }

  return {
    mode: "chat",
  };
}
