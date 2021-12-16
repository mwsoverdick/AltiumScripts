{.................................................................................}
{ Summary   This Script verifies that all components that are not fitted in       }
{           a specific "template" variant are also not fitted in all other        }
{           variants. This is used to ensure that all permenant DNP components    }
{           do not end up on production assemblies.                               ]
{                                                                                 }
{ Note      This script does not destinguish between "Not Fitted" and alt part    }
{           in the template. Any "Component Variation" that is not "Fitted" in    }
{           the template, will be "Not Fitted" in all other variants.             }
{                                                                                 }
{ Warning   This script can only set components to "Not Fitted" it cannot undo    }
{           a change if a component is no longer a permenant DNP. This must be    }
{           done manually to ensure accuracy.                                     }
{                                                                                 }
{ Created by:    Mitchell Overdick 20/03/2021                                     }
{ Copyright (C) 2021 Mitchell Overdick                                            }
{                                                                                 }
{ This program is free software: you can redistribute it and/or modify it under   }
{ the terms of the GNU General Public License as published by the Free Software   }
{ Foundation, either version 3 of the License, or (at your option) any later      }
{ version.                                                                        }
{                                                                                 }
{ This program is distributed in the hope that it will be useful, but WITHOUT ANY }
{ WARRANTY; without even the implied warranty of  MERCHANTABILITY or FITNESS FOR  }
{ A PARTICULAR PURPOSE. See the GNU General Public License for more details.      }
{                                                                                 }
{ You should have received a copy of the GNU General Public License along with    }
{ this program.  If not, see <http://www.gnu.org/licenses/>.                      }
{.................................................................................}

{.................................................................................}
procedure TemplateCompare;
var
    Workspace             : IWorkspace;
    PCBProject            : IProject;
    ComponentNum          : Integer;
    VariantNum            : Integer;
    Variant               : IProjectVariant;

    CompUid               : WideString;
    CompVariation         : IComponentVariation;
    CompVariationBase     : IComponentVariation;

    TemplateVariantName   : String;
    TemplateVariant       : IProjectVariant;
    UpdatedCount          : Integer;

begin
    TemplateVariantName := 'Engineering';    // Name of variant to use as template
    TemplateVariant     := Nil;
    Workspace           := GetWorkspace;
    PCBProject          := Workspace.DM_FocusedProject;

    // Can only be run on PCB projects
    if (PcbProject = nil) then
    begin
        ShowError('Current Project is not a PCB Project');
        exit;
    end;

    if (AnsiUpperCase(ExtractFileExt(PCBProject.DM_ProjectFileName)) <> '.PRJPCB') then
    begin
        ShowError('Current Project is not a PCB Project');
        exit;
    end;

    // Search for base variant
    for VariantNum := 0 to PCBProject.DM_ProjectVariantCount - 1 do
    begin
        Variant := PCBProject.DM_ProjectVariants(VariantNum);   // Get the variant object

        // Store template variant for later
        if (Variant.DM_Description = TemplateVariantName) then
        begin
            TemplateVariant := Variant;
        end;
    end;

    // Ensure template variant was found
    if (TemplateVariant = Nil) then
    begin
        ShowError('No variant named "' + TemplateVariantName + '" found!');
        exit;
    end;

    // Check every component in template variation
    UpdatedCount := 0;
    for ComponentNum := 0 to TemplateVariant.DM_VariationCount - 1 do
    begin

        // Look up the component's variation in every variant
        for VariantNum := 0 to PCBProject.DM_ProjectVariantCount - 1 do
        begin

            Variant := PCBProject.DM_ProjectVariants(VariantNum);

            // Only check variation if current variant is not the template
            if (Variant.DM_Description <> TemplateVariantName) then
            begin

                CompUid := TemplateVariant.DM_Variations(ComponentNum).DM_UniqueId;
                CompVariation := Variant.DM_FindComponentVariationByUniqueId(CompUid);

                // If the component variation is Nil, we need to create it first
                if (CompVariation = Nil) then
                begin
                    CompVariation := Variant.DM_AddComponentVariation();
                    CompVariation.DM_SetUniqueId(CompUid);
                end;

                // Set to "Not Fitted" if not already
                if CompVariation.DM_VariationKind <> eVariation_NotFitted then
                begin
                    CompVariation.DM_SetVariationKind(eVariation_NotFitted);
                    UpdatedCount := UpdatedCount + 1;
                end;
            end;
        end;
    end;

    // Provide user with report
    ShowInfo(IntToStr(TemplateVariant.DM_VariationCount) + ' components checked.'
             + sLineBreak +
             IntToStr(UpdatedCount) + ' changes made.'
             + sLineBreak + sLineBreak +
             'Manually verify to ensure correctness.');
end;
