namespace VMS.TPS

open System
open System.IO
open VMS.TPS.Common.Model.API
open VMS.TPS.Common.Model.Types
open System.Reflection

module App =

    /////////////////////
    ///// Filenames /////
    /////////////////////
    let directory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\"
    let fullList = "PlanIDList.txt"
    let idListFilename = directory + fullList
    let resultsFilename = directory + "results.csv"
    
    //////////////////////////
    ///// Structure info /////
    //////////////////////////
    let givenContours = [|
        "1 CTV - 50.4_P";
        "1 CTV - 79.2_P";
        "2 PTV - 50.4_P";
        "2 PTV - 79.2_P";
        "Bladder_P";
        "BODY_P";
        "Bone_Femur_L_P";
        "Bone_Femur_R_P";
        "Bone_Iliac_P";
        "Bone_PubicSymph_";
        "Bone_Sacrum_P";
        "Bone_Spine_P";
        "Bowel_Bag_P";
        "Bowel_Large_P";
        "Bowel_Sigmoid_P";
        "Bowel_Small_P";
        "CouchInterior_P";
        "CouchSurface_P";
        "PenileBulb_P";
        "Prostate_P";
        "Rectum_P";
        "SeminalVes_Full_";
        "SeminalVes_Prox_";
        "SpaceOAR_P";
        "POI_1"
    |]

    let givenTargets = [|
        "1 CTV - 50.4_P";
        "1 CTV - 79.2_P";
        "2 PTV - 50.4_P";
        "2 PTV - 79.2_P"
    |]

    let givenOARs = Set.difference (Set.ofArray givenContours) (Set.ofArray givenTargets) |> Set.toArray

    /////////////////////////////////////////////////////////////////////////////////////////////
    //////////////////////////////////////Add New Tests Here/////////////////////////////////////
    /////////////////////////////////////////////////////////////////////////////////////////////
    type TestNames =
    | AssignedNumber = 0
    | CourseID = 1
    | PlanID = 2
    | PlanApprovalStatus = 3
    | PlanTarget = 4
    | PlanTechnique = 5
    | NumberOfBeams = 6
    | EnergiesUsed = 7
    | PlanMUs = 8
    | HotspotLocation = 9
    | CreatedStructures = 10
    | TargetD95 = 11
    | TargetD99 = 12
    | TargetD003cc = 13
    | BladderD15 = 14
    | BladderD30 = 15
    | BladderD60 = 16
    | RectumD10 = 17
    | RectumD20 = 18
    | RectumD40 = 19
    | Femur_LD10 = 20
    | Femur_RD10 = 21
    | NTO = 22
    | MCO = 23
    | gEUD = 24
    | OptimizationStructures = 25

    type TestResult = {
        TestName: string
        TestResult: string
    }

    let Tests : TestNames[] = Enum.GetValues(typeof<TestNames>) :?> TestNames[]

    ////////////////////////////
    ///// Helper Functions /////
    ////////////////////////////
    let formatProgressMessage msg id =
        (sprintf "\rEvaluating %s: %s" id msg).PadRight(70)

    let printProgressMessage msg id =
        Console.Write(formatProgressMessage msg id)

    let printProgressMessageWithNewline msg id =
        Console.WriteLine(formatProgressMessage msg id)

    let wrapInQuotes str =
        "\"" + str + "\""

    let getPatientId (plan:PlanSetup) =
        plan.Course.Patient.Id

    let getClosestStructure structureName (plan:PlanSetup) =
        Seq.cast<Structure> plan.StructureSet.Structures |>
        Seq.filter(fun x -> x.Id.Contains(structureName)) |>
        Seq.sortBy(fun x -> x.Id.Length) |>
        Seq.head

    let getDoseToVolume structureName volume volumePresentation requestedDosePresentation (plan:PlanSetup) =
        let structure = plan |> getClosestStructure structureName
        plan.GetDoseAtVolume(structure, volume, volumePresentation, requestedDosePresentation)
        
    let ``getDose%ToVolume%`` structureName volume (plan:PlanSetup) =
        plan |> getDoseToVolume structureName volume VolumePresentation.Relative DoseValuePresentation.Relative
                
    let ``getDose%ToVolume(cc)`` structureName volume (plan:PlanSetup) =
        plan |> getDoseToVolume structureName volume VolumePresentation.AbsoluteCm3 DoseValuePresentation.Relative
                
    let ``getDose(cGy)ToVolume%`` structureName volume (plan:PlanSetup) =
        plan |> getDoseToVolume structureName volume VolumePresentation.Relative DoseValuePresentation.Absolute
                
    let ``getDose(cGy)ToVolume(cc)`` structureName volume (plan:PlanSetup) =
        plan |> getDoseToVolume structureName volume VolumePresentation.AbsoluteCm3 DoseValuePresentation.Absolute

    let getVolumeAtDose structureName dose requestedVolumePresentation (plan:PlanSetup) =
        let structure = plan |> getClosestStructure structureName
        plan.GetVolumeAtDose(structure, dose, requestedVolumePresentation)

    let ``getVolume%AtDose`` structureName dose (plan:PlanSetup) =
        plan |> getVolumeAtDose structureName dose VolumePresentation.Relative
        
    let ``getVolume(cc)AtDose`` structureName dose (plan:PlanSetup) =
        plan |> getVolumeAtDose structureName dose VolumePresentation.AbsoluteCm3
        
        
    /////////////////////////////////////
    ///// Individual test functions /////
    /////////////////////////////////////

    let getAssignedNumber (plan:PlanSetup) = 
        {
            TestName = "Assigned Number";
            TestResult = plan.Course.Patient.FirstName
        }

    let getCourseId (plan:PlanSetup) = 
        {
            TestName = "Course ID"
            TestResult = plan.Course.Id
        }

    let getPlanId (plan:PlanSetup) = 
        {
            TestName = "Plan ID"
            TestResult = plan.Id
        }
    
    let getPlanApprovalStatus (plan:PlanSetup) = 
        getPatientId plan |> printProgressMessage "Approval Status..."
        {
            TestName = "Plan Approval Status"
            TestResult = plan.ApprovalStatus.ToString()
        }
        
    let getPlanTarget (plan:PlanSetup) =
        getPatientId plan |> printProgressMessage "Target..."
        {
            TestName = "Target Structure"
            TestResult = plan.TargetVolumeID
        }
        
                
    let getPlanTechnique (plan:PlanSetup) =
        getPatientId plan |> printProgressMessage "Plan Technique..."
        {
            TestName = "Plan Technique"
            TestResult = 
                let (|IMRT|FiF|VMAT|DCA|Static|Unknown|) (beam:Beam) =
                    match beam.MLCPlanType with
                    | MLCPlanType.VMAT -> VMAT
                    | MLCPlanType.ArcDynamic -> DCA
                    | MLCPlanType.Static -> Static
                    | MLCPlanType.DoseDynamic ->
                        if beam.ControlPoints.Count > 20 then
                            IMRT
                        else
                            FiF
                    | _ -> Unknown

                match (Seq.cast<Beam> plan.Beams |> Seq.filter (fun x -> not x.IsSetupField) |> Seq.head) with
                | VMAT -> "VMAT"
                | IMRT -> "IMRT"
                | DCA -> "DCA"
                | FiF -> "FiF"
                | Static -> "Static"
                | Unknown -> "Unknown"
        }
                
    let getNumberOfBeams (plan:PlanSetup) =
        getPatientId plan |> printProgressMessage "Number of Beams..."
        {
            TestName = "# of Beams"
            TestResult =
                Seq.cast<Beam> plan.Beams |>
                Seq.filter (fun x -> not x.IsSetupField) |>
                Seq.length |>
                sprintf "%i"
        }
                
    let getEnergies (plan:PlanSetup) =
        getPatientId plan |> printProgressMessage "Energies..."
        {
            TestName = "Energies Used"
            TestResult =
                Seq.cast<Beam> plan.Beams |>
                Seq.filter (fun x -> not x.IsSetupField) |>
                Seq.map (fun x -> x.EnergyModeDisplayName) |>
                Seq.distinct |>
                String.concat "\n" |>
                wrapInQuotes
        }
        
    let getPlanMUs (plan:PlanSetup) =
        getPatientId plan |> printProgressMessage "MUs..."
        {
            TestName = "Plan MUs"
            TestResult =
                Seq.cast<Beam> plan.Beams |>
                Seq.filter (fun x -> not x.IsSetupField) |>
                Seq.map (fun x -> x.Meterset.Value) |>
                Seq.fold (fun state x -> state + x) 0.0 |>
                sprintf "%.1f"
        }

    let getDoseConstraint structureName doseOrVolume number firstUnits secondUnits (plan:PlanSetup) =
        let constraintText = sprintf "%s %s%.2f%s[%s]" structureName doseOrVolume number firstUnits secondUnits
        getPatientId plan |> printProgressMessage (sprintf "%s..." constraintText)
        let constraintError = sprintf "%s is not a recognized constraint" constraintText
        {
            TestName = constraintText
            TestResult =
                match doseOrVolume with
                | "D" -> 
                    match firstUnits, secondUnits with
                    | "%", "%" -> (plan |> ``getDose%ToVolume%`` structureName number).ToString()
                    | "%", "cGy" -> (plan |> ``getDose(cGy)ToVolume%`` structureName number).ToString()
                    | "%", "Gy" -> (plan |> ``getDose(cGy)ToVolume%`` structureName number).ToString()
                    | "cc", "%" -> (plan |> ``getDose%ToVolume(cc)`` structureName number).ToString()
                    | "cc", "cGy" -> (plan |> ``getDose(cGy)ToVolume(cc)`` structureName number).ToString()
                    | "cc", "Gy" -> (plan |> ``getDose(cGy)ToVolume(cc)`` structureName number).ToString()
                    | _, _ -> constraintError
                | "V" -> 
                    match firstUnits, secondUnits with
                    | "%", "%" -> (plan |> ``getVolume%AtDose`` structureName (new DoseValue(number, DoseValue.DoseUnit.Percent))).ToString()
                    | "cGy", "%" -> (plan |> ``getVolume%AtDose`` structureName (new DoseValue(number, DoseValue.DoseUnit.cGy))).ToString()
                    | "Gy", "%" -> (plan |> ``getVolume%AtDose`` structureName (new DoseValue(number, DoseValue.DoseUnit.Gy))).ToString()
                    | "%", "cc" -> (plan |> ``getVolume(cc)AtDose`` structureName (new DoseValue(number, DoseValue.DoseUnit.Percent))).ToString()
                    | "cGy", "cc" -> (plan |> ``getVolume(cc)AtDose`` structureName (new DoseValue(number, DoseValue.DoseUnit.cGy))).ToString()
                    | "Gy", "cc" -> (plan |> ``getVolume(cc)AtDose`` structureName (new DoseValue(number, DoseValue.DoseUnit.Gy))).ToString()
                    | _, _ -> constraintError
                | _ -> constraintError
        }

    // Returns list of structure IDs that the hotspot is within
    // Only takes the smallest of the structures from the "givenTargets" list and any other of the "givenContours" that are not BODY
    let getHotspot (plan:PlanSetup) =
        getPatientId plan |> printProgressMessage "Hotspot Location..."
        {
            TestName = "Hotspot Location"
            TestResult =
                let hotspotLocation = plan.Dose.DoseMax3DLocation

                // Preexisting structures that are contoured (minus the Body structure)
                let existingContouredStructures = 
                    Seq.cast<Structure> plan.StructureSet.Structures |>
                    Seq.filter (fun x -> x.HasSegment) |>
                    Seq.filter (fun x -> givenContours |> Array.contains x.Id) |>
                    Seq.filter (fun x -> not (x.Id.ToLower().Contains("body")))

                // List of structures that the hotspot is within
                let hotspotStructureList = 
                    existingContouredStructures |>
                    Seq.filter (fun x -> x.IsPointInsideSegment hotspotLocation) |>
                    Seq.sortBy(fun x -> x.Volume)

        
                Seq.append 
                    // Only the smallest structure from the targets list
                    (match hotspotStructureList |> 
                            Seq.filter (fun x -> givenTargets |> Array.contains x.Id) |> 
                            Seq.tryHead with
                            | Some x -> Seq.singleton x
                            | None -> Seq.empty)
                    // And then everything else from the OARs list
                    (hotspotStructureList |> 
                    Seq.filter (fun x -> givenOARs |> Array.contains x.Id)) |>
                Seq.map(fun x -> x.Id) |>
                String.concat "\n" |>
                wrapInQuotes
        }

    let getCreatedStructures (plan:PlanSetup) =
        getPatientId plan |> printProgressMessage "Created Structures..."
        {
            TestName = "Created Structures"
            TestResult = 
                Seq.cast<Structure> plan.StructureSet.Structures |>
                Seq.filter (fun x -> not (Array.contains x.Id givenContours)) |>
                Seq.map (fun x -> x.Id) |>
                String.concat "\n" |>
                wrapInQuotes
        }
        
    let getNTOAutoOrManual (plan:PlanSetup) =
        getPatientId plan |> printProgressMessage "NTO..."

        let getManualParameters =
            Seq.cast<OptimizationParameter> plan.OptimizationSetup.Parameters |>
            Seq.filter (fun x -> x.GetType() = typeof<OptimizationNormalTissueParameter>) |>
            Seq.map (fun x -> 
                let y = x :?> OptimizationNormalTissueParameter
                sprintf "Manual\nPriority: %.0f\nDistance from Target Border (mm): %.1f\nFalloff: %.2f\nStart Dose (%%): %.1f\nEnd Dose (%%): %.1f" y.Priority y.DistanceFromTargetBorderInMM y.FallOff y.StartDosePercentage y.EndDosePercentage
                ) |>
            Seq.head

        {
            TestName = "NTO"
            TestResult =
                Seq.cast<OptimizationParameter> plan.OptimizationSetup.Parameters |>
                Seq.filter (fun x -> x.GetType() = typeof<OptimizationNormalTissueParameter>) |>
                Seq.map (fun x -> 
                            match (x :?> OptimizationNormalTissueParameter).IsAutomatic with 
                            | true -> sprintf "Auto\nPriority: %.0f" (x :?> OptimizationNormalTissueParameter).Priority
                            | false -> getManualParameters) |>
                Seq.head |>
                wrapInQuotes
        }

    let getMCOUseage (plan:PlanSetup) =
        getPatientId plan |> printProgressMessage "MCO..."
        {
            TestName = "MCO"
            TestResult = 
                match (plan :?> ExternalPlanSetup).TradeoffExplorationContext.CanCreatePlanCollection with 
                | true -> "Yes"
                | false -> "No"
        }

    let getGEUDUseage (plan:PlanSetup) =
        getPatientId plan |> printProgressMessage "gEUD..."

        let numOfgEUDObjectives = 
            Seq.cast<OptimizationObjective> plan.OptimizationSetup.Objectives |>
            Seq.filter (fun x -> x.GetType() = typeof<OptimizationEUDObjective>) |>
            Seq.length

        {
            TestName = "gEUD"
            TestResult =
                match numOfgEUDObjectives with 
                | 0 -> "No"
                | _ -> "Yes"
        }
        

    let getOptimizationStructures (plan:PlanSetup) =
        getPatientId plan |> printProgressMessage "Optimization Structures..."

        let eudObjectives = 
            Seq.cast<OptimizationObjective> plan.OptimizationSetup.Objectives |>
            Seq.filter (fun x -> x.GetType() = typeof<OptimizationEUDObjective>) |>
            Seq.cast<OptimizationEUDObjective> |>
            Seq.map (fun x -> sprintf "gEUD Objectives:\n%s %A %A with a = %.2f (Priority = %.0f)" x.StructureId x.Operator x.Dose x.ParameterA x.Priority) |>
            String.concat "\n"
            
        let lineObjectives = 
            Seq.cast<OptimizationObjective> plan.OptimizationSetup.Objectives |>
            Seq.filter (fun x -> x.GetType() = typeof<OptimizationLineObjective>) |>
            Seq.cast<OptimizationLineObjective> |>
            Seq.map (fun x -> sprintf "Line Objectives:\n%s %A (Priority = %.0f)" x.StructureId x.Operator x.Priority)|>
                       String.concat "\n"
            
        let meanDoseObjectives = 
            Seq.cast<OptimizationObjective> plan.OptimizationSetup.Objectives |>
            Seq.filter (fun x -> x.GetType() = typeof<OptimizationMeanDoseObjective>) |>
            Seq.cast<OptimizationMeanDoseObjective> |>
            Seq.map (fun x -> sprintf "Mean Dose Objectives:\n%s %A %A (Priority = %.0f)" x.StructureId x.Operator x.Dose x.Priority)|>
                       String.concat "\n"
            
        let pointObjectives = 
            Seq.cast<OptimizationObjective> plan.OptimizationSetup.Objectives |>
            Seq.filter (fun x -> x.GetType() = typeof<OptimizationPointObjective>) |>
            Seq.cast<OptimizationPointObjective> |>
            Seq.map (fun x -> sprintf "Point Objectives:\n%s %A %A to %.1f%% (Priority = %.0f)" x.StructureId x.Operator x.Dose x.Volume x.Priority)|>
                       String.concat "\n"
        {
            TestName = "Optimization Structures"
            TestResult = (sprintf "%s\n\n%s\n\n%s\n\n%s" pointObjectives meanDoseObjectives eudObjectives lineObjectives) |> wrapInQuotes
        }        

    /////////////////////////
    ///// Print results /////
    /////////////////////////
    let toCSV (results:(TestResult[])[]) =
        let outResults = 
            results |>
            Array.map(fun result -> 
                result |>
                            Array.map (fun x -> x.TestResult) |>
                            String.concat ","
            )

        Console.WriteLine ("\rDone!                                              ")
        let testNames = results |>
                        Array.head |>
                        Array.map (fun x -> x.TestName) |>
                        String.concat ","

        File.WriteAllLines(resultsFilename, Array.append (Array.singleton testNames) outResults)
        Console.WriteLine (sprintf "Results have been saved to %s" resultsFilename)
       
    //////////////////////
    ///// Validation /////
    //////////////////////
    let validateCourse (courseId:string) (patient:Patient) =
        let lowerCourse = courseId.ToLower()

        let validCourses = 
            Seq.cast<Course> patient.Courses |>
            Seq.filter(fun x -> x.Id.ToLower().Contains(lowerCourse))
        
        if Seq.length validCourses <> 1 then
            None
        else
            Some (Seq.exactlyOne validCourses)
        
    let validatePlans (course:Course) =
        let validPlans =
            Seq.cast<PlanSetup> course.PlanSetups |>
            Seq.filter(fun x -> x.ApprovalStatus = PlanSetupApprovalStatus.Reviewed || x.ApprovalStatus = PlanSetupApprovalStatus.PlanningApproved || x.ApprovalStatus = PlanSetupApprovalStatus.TreatmentApproved) |>
            Seq.toArray
        
        if Seq.length validPlans = 0 then
            None
        else
            Some validPlans
        
    let validateDoses (plans:PlanSetup [] option) =
        match plans with
        | None -> Array.empty
        | Some list-> list |> Array.map(fun x -> 
            match x.IsDoseValid with
            | true -> Some x
            | false -> None)

    let getPlans course (patient:Patient) =
        patient |>
        validateCourse course |>
        Option.bind validatePlans |>
        validateDoses |> 
        Array.filter (fun x -> x.IsSome) |>
        Array.map (fun x -> x.Value)

    ////////////////////////////////
    ///// Main Script Function /////
    ////////////////////////////////
    let exectueScript (app:Application) =         
        let idList = File.ReadAllLines idListFilename
        
        let results = 
            idList |> Array.map(fun id -> 
                id |> printProgressMessage "Opening Patient..."
        
                let patient = app.OpenPatientById(id)
        
                /////////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////////Add New Tests Here/////////////////////////////////////
                /////////////////////////////////////////////////////////////////////////////////////////////
                let runTests plan = 
                    Tests |> Array.map (fun test ->
                        match test with
                        | TestNames.AssignedNumber -> getAssignedNumber plan
                        | TestNames.CourseID -> getCourseId plan
                        | TestNames.PlanID -> getPlanId plan
                        | TestNames.PlanApprovalStatus -> getPlanApprovalStatus plan
                        | TestNames.PlanTarget -> getPlanTarget plan
                        | TestNames.PlanTechnique -> getPlanTechnique plan
                        | TestNames.NumberOfBeams -> getNumberOfBeams plan
                        | TestNames.EnergiesUsed -> getEnergies plan
                        | TestNames.PlanMUs -> getPlanMUs plan
                        | TestNames.HotspotLocation -> getHotspot plan
                        | TestNames.CreatedStructures -> getCreatedStructures plan
                        | TestNames.TargetD95 -> getDoseConstraint "PTV - 79.2" "D" 95.0 "%" "%" plan
                        | TestNames.TargetD99 -> getDoseConstraint "PTV - 79.2" "D" 99.0 "%" "%" plan
                        | TestNames.TargetD003cc -> getDoseConstraint "PTV - 79.2" "D" 0.03 "cc" "%" plan
                        | TestNames.BladderD15 -> getDoseConstraint "Bladder" "D" 15.0 "%" "%" plan
                        | TestNames.BladderD30 -> getDoseConstraint "Bladder" "D" 30.0 "%" "%" plan
                        | TestNames.BladderD60 -> getDoseConstraint "Bladder" "D" 60.0 "%" "%" plan
                        | TestNames.RectumD10 -> getDoseConstraint "Rectum" "D" 10.0 "%" "%" plan
                        | TestNames.RectumD20 -> getDoseConstraint "Rectum" "D" 20.0 "%" "%" plan
                        | TestNames.RectumD40 -> getDoseConstraint "Rectum" "D" 40.0 "%" "%" plan
                        | TestNames.Femur_LD10 -> getDoseConstraint "Femur_L" "D" 10.0 "%" "%" plan
                        | TestNames.Femur_RD10 -> getDoseConstraint "Femur_R" "D" 10.0 "%" "%" plan
                        | TestNames.NTO -> getNTOAutoOrManual plan
                        | TestNames.MCO -> getMCOUseage plan
                        | TestNames.gEUD -> getGEUDUseage plan
                        | TestNames.OptimizationStructures -> getOptimizationStructures plan
                    )
        
                let result = getPlans "Prost" patient |> Array.map(fun x -> x |> runTests)
        
                app.ClosePatient()
                        
                id |> printProgressMessageWithNewline "Done"
        
                result
            )

        let flattenedResults = results |> Array.reduce Array.append
        
        try
            do toCSV flattenedResults
        with e -> Console.WriteLine(e.Message)

        app.ClosePatient()

    //////////////////////////////
    ///// Script entry point /////
    //////////////////////////////
    [<EntryPoint; STAThread>]
    let main _ =
        Console.Write("Logging into Eclipse...")
        
        let app = Application.CreateApplication()

        do exectueScript app
        
        app.Dispose()
        
        Console.ReadKey() |> ignore
        
        1

        