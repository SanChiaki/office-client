export type RouteDecision =
  | {
      mode: "skill";
      skillName: string;
    }
  | {
      mode: "chat";
    };

export function decideRoute(input: string): RouteDecision {
  const trimmed = input.trim();

  if (/^\/upload_data(?:$|\s)/.test(trimmed)) {
    return {
      mode: "skill",
      skillName: "upload_data",
    };
  }

  return {
    mode: "chat",
  };
}
