"use client";

import * as React from "react";
import { formatDateTime } from "@/lib/date";
import { formatQuantity } from "@/lib/display";
import { useDemoSession } from "@/lib/session-context";
import { EmptyState, PageHeader, PermissionNotice, SectionTitle, Surface } from "@/components/ui";

export default function RecipesPage() {
  const { activeHousehold, permissions, state, visibleInventory, actions } = useDemoSession();
  const householdRecipes = state.recipes.filter((entry) => entry.householdId === activeHousehold.id);
  const [selectedRecipeId, setSelectedRecipeId] = React.useState<string | null>(householdRecipes[0]?.id ?? null);
  const [servings, setServings] = React.useState(householdRecipes[0]?.servings ?? 1);
  const [importUrl, setImportUrl] = React.useState("https://example.com/weeknight-pantry-chili");

  React.useEffect(() => {
    if (!householdRecipes.length) {
      setSelectedRecipeId(null);
      return;
    }

    if (!householdRecipes.some((entry) => entry.id === selectedRecipeId)) {
      setSelectedRecipeId(householdRecipes[0].id);
      setServings(householdRecipes[0].servings);
    }
  }, [householdRecipes, selectedRecipeId]);

  const selectedRecipe = householdRecipes.find((entry) => entry.id === selectedRecipeId) ?? null;
  const selectedIngredients = selectedRecipe
    ? state.recipeIngredients.filter((entry) => entry.recipeId === selectedRecipe.id)
    : [];
  const relevantLogs = state.recipeLogs
    .filter((entry) => entry.householdId === activeHousehold.id)
    .sort((left, right) => new Date(right.cookedAt).getTime() - new Date(left.cookedAt).getTime());

  return (
    <div className="page">
      <PageHeader
        eyebrow="Recipe engine"
        title="Turn meals into inventory events."
        description="Recipes stay grounded in the same product catalog as inventory, so cooking a meal deducts ingredients and changes tomorrow's forecast."
      />

      {!permissions.canCookRecipes ? (
        <PermissionNotice>
          Contributor mode can browse recipes, but only owners and members can log a cook and deduct ingredients.
        </PermissionNotice>
      ) : null}

      <div className="twoColumn">
        <Surface strong>
          <SectionTitle title="Recipe library" action={<span className="badge badgeOrange">{householdRecipes.length} recipes</span>} />
          <div className="stack">
            {householdRecipes.map((recipe) => (
              <button
                key={recipe.id}
                type="button"
                className="listItem"
                onClick={() => {
                  setSelectedRecipeId(recipe.id);
                  setServings(recipe.servings);
                }}
                style={{
                  textAlign: "left",
                  borderColor: recipe.id === selectedRecipeId ? "rgba(201, 79, 21, 0.4)" : undefined,
                  background: recipe.id === selectedRecipeId ? "rgba(255, 241, 223, 0.88)" : undefined,
                }}
              >
                <strong>{recipe.name}</strong>
                <p className="muted">{recipe.instructions}</p>
              </button>
            ))}
          </div>
        </Surface>

        <Surface>
          {selectedRecipe ? (
            <div className="stack">
              <SectionTitle title={selectedRecipe.name} action={<span className="badge badgeGreen">{selectedRecipe.servings} base servings</span>} />
              <p className="muted">{selectedRecipe.instructions}</p>
              <div className="cluster">
                <label className="field" style={{ maxWidth: "180px" }}>
                  <span className="label">Servings to log</span>
                  <input className="input" type="number" min="1" value={servings} onChange={(event) => setServings(Number(event.target.value))} />
                </label>
                <button className="button" type="button" onClick={() => actions.cookRecipe(selectedRecipe.id, servings)} disabled={!permissions.canCookRecipes}>
                  Log cook + deduct inventory
                </button>
              </div>
              <div className="stack">
                {selectedIngredients.map((ingredient) => {
                  const product = state.products.find((entry) => entry.id === ingredient.productId);
                  const inventoryMatch = visibleInventory.find((entry) => entry.productId === ingredient.productId);
                  const scaledAmount = (ingredient.quantity * servings) / Math.max(selectedRecipe.servings, 1);
                  return (
                    <div key={ingredient.id} className="listItem">
                      <div className="sectionHeading">
                        <strong>{product?.name ?? "Unknown ingredient"}</strong>
                        <span className="badge badgeNeutral">{formatQuantity(scaledAmount, ingredient.unit)}</span>
                      </div>
                      <p className="muted">
                        Inventory on hand:{" "}
                        {inventoryMatch ? formatQuantity(inventoryMatch.quantityValue, inventoryMatch.quantityUnit) : "Not currently stocked"}
                      </p>
                    </div>
                  );
                })}
              </div>
            </div>
          ) : (
            <EmptyState title="No recipe selected" description="Pick a recipe from the library to preview ingredient deductions." />
          )}
        </Surface>
      </div>

      <div className="twoColumn">
        <Surface>
          <SectionTitle title="Recipe URL import" action={<span className="badge badgeNeutral">Simulated parse</span>} />
          <div className="stack">
            <label className="field">
              <span className="label">Recipe URL</span>
              <input className="input" value={importUrl} onChange={(event) => setImportUrl(event.target.value)} />
            </label>
            <button className="buttonSecondary" type="button" onClick={() => actions.importRecipeFromUrl(importUrl)} disabled={!permissions.canCookRecipes}>
              Simulate import
            </button>
            <p className="muted">This prototype adds a structured draft recipe with ingredients that map into the same product catalog and inventory logic.</p>
          </div>
        </Surface>

        <Surface>
          <SectionTitle title="Recent cook logs" />
          {relevantLogs.length ? (
            <ul className="list">
              {relevantLogs.map((entry) => {
                const recipe = state.recipes.find((item) => item.id === entry.recipeId);
                return (
                  <li key={entry.id} className="listItem">
                    <strong>{recipe?.name ?? "Recipe"}</strong>
                    <p className="muted">{entry.servingsMade} servings logged</p>
                    <span className="muted mono">{formatDateTime(entry.cookedAt)}</span>
                  </li>
                );
              })}
            </ul>
          ) : (
            <EmptyState title="No cook logs yet" description="Log one of the household recipes to see inventory deductions land here." />
          )}
        </Surface>
      </div>
    </div>
  );
}
