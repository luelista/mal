-- This started out as a simple public variant record.
-- Then smart pointers were added.  They were part of the Mal_Type and
-- were required to be public because of the dependencies and
-- how the variant record was public.  Not very Ada-like.
-- The third version bites the bullet and delares Mal_Type as tagged.
-- Smart pointers are an OO version in a separate package.
-- The Doubly_Linked_Lists have been replaced with a tree-like list instead...

-- WARNING!  This code contains:
-- Recursive data structures.
-- Object-based smart pointers.
-- Object-oriented code.
-- And strong-typing!

-- Chris M Moore 25/03/2015

with Ada.Strings.Unbounded;
with Smart_Pointers;

package Types is

   -- Some simple types.  Not supposed to use the standard types directly.

   subtype Mal_Float is Float;
   subtype Mal_Integer is Integer;
   subtype Mal_String is String;

   -- Start off with the top-level abstract type.

   subtype Mal_Handle is Smart_Pointers.Smart_Pointer;

   type Sym_Types is (Int, Floating, List, Sym, Str, Atom,
                      Unitary, Node, Lambda, Error);

   type Mal_Type is abstract new Smart_Pointers.Base_Class with private;

   function Sym_Type (T : Mal_Type) return Sym_Types is abstract;

   function Get_Meta (T : Mal_Type) return Mal_Handle;

   procedure Set_Meta (T : in out Mal_Type'Class; SP : Mal_Handle);

   function To_String (T : Mal_Type'Class) return Mal_String;

   type Mal_Ptr is access all Mal_Type'Class;

   -- A helper function that just view converts the smart pointer to
   -- a Mal_Type'Class pointer.
   function Deref (S : Mal_Handle) return Mal_Ptr;

   -- A helper function to detect null smart pointers.
   function Is_Null (S : Mal_Handle) return Boolean;

   -- Derived types.  All boilerplate from here.

   type Int_Mal_Type is new Mal_Type with private;

   function New_Int_Mal_Type (Int : Mal_Integer) return Mal_Handle;

   overriding function Sym_Type (T : Int_Mal_Type) return Sym_Types;

   function Get_Int_Val (T : Int_Mal_Type) return Mal_Integer;

   type Int_Ptr is access all Int_Mal_Type;

   function Deref_Int (SP : Mal_Handle) return Int_Ptr;


   type Float_Mal_Type is new Mal_Type with private;

   function New_Float_Mal_Type (Floating : Mal_Float) return Mal_Handle;

   overriding function Sym_Type (T : Float_Mal_Type) return Sym_Types;

   function Get_Float_Val (T : Float_Mal_Type) return Mal_Float;

   type Float_Ptr is access all Float_Mal_Type;

   function Deref_Float (SP : Mal_Handle) return Float_Ptr;


   type Sym_Mal_Type is new Mal_Type with private;

   function New_Sym_Mal_Type (Sym : Character) return Mal_Handle;

   overriding function Sym_Type (T : Sym_Mal_Type) return Sym_Types;

   function Symbol (T : Sym_Mal_Type) return Character;

   type Sym_Ptr is access all Sym_Mal_Type;

   function Deref_Sym (S : Mal_Handle) return Sym_Ptr;


   type String_Mal_Type is new Mal_Type with private;

   function New_String_Mal_Type (Str : Mal_String) return Mal_Handle;

   overriding function Sym_Type (T : String_Mal_Type) return Sym_Types;

   function Get_String (T : String_Mal_Type) return Mal_String;


   type Atom_Mal_Type is new Mal_Type with private;

   function New_Atom_Mal_Type (Str : Mal_String) return Mal_Handle;

   overriding function Sym_Type (T : Atom_Mal_Type) return Sym_Types;

   function Get_Atom (T : Atom_Mal_Type) return Mal_String;

   type Atom_Ptr is access all Atom_Mal_Type;

   function Deref_Atom (S : Mal_Handle) return Atom_Ptr;



   type Error_Mal_Type is new Mal_Type with private;

   function New_Error_Mal_Type (Str : Mal_String) return Mal_Handle;

   overriding function Sym_Type (T : Error_Mal_Type) return Sym_Types;


   type Unitary_Functions is
     (Quote, Unquote, Quasiquote, Splice_Unquote, Deref);

   type Unitary_Mal_Type is new Mal_Type with private;

   function New_Unitary_Mal_Type (Func : Unitary_Functions; Op : Mal_Handle)
   return Mal_Handle;

   overriding function Sym_Type (T : Unitary_Mal_Type) return Sym_Types;

   function Get_Func (T : Unitary_Mal_Type) return Unitary_Functions;

   function Get_Op (T : Unitary_Mal_Type) return Mal_Handle;


   -- Lists.

   type List_Types is (List_List, Vector_List, Hashed_List);
   function Opening (LT : List_Types) return Character;
   function Closing (LT : List_Types) return Character;

   type List_Mal_Type is new Mal_Type with private;

   function New_List_Mal_Type
     (List_Type : List_Types;
      The_First_Node : Mal_Handle := Smart_Pointers.Null_Smart_Pointer)
   return Mal_Handle;

   overriding function Sym_Type (T : List_Mal_Type) return Sym_Types;

   function Get_List_Type (L : List_Mal_Type) return List_Types;

   procedure Append (To_List : in out List_Mal_Type; Op : Mal_Handle);

   function Length (L : List_Mal_Type) return Natural;

   -- Get the first item in the list:
   function Car (L : List_Mal_Type) return Mal_Handle;

   -- Get the rest of the list (second item onwards)
   function Cdr (L : List_Mal_Type) return Mal_Handle;

   type Func_Access is access
     function (Elem : Mal_Handle)
     return Mal_Handle;

   function Map
     (Func_Ptr : Func_Access;
      L : List_Mal_Type)
   return Mal_Handle;

   type Binary_Func_Access is access
     function (A, B : Mal_Handle)
     return Mal_Handle;

   function Reduce
     (Func_Ptr : Binary_Func_Access;
      L : List_Mal_Type)
   return Mal_Handle;

   function Null_List (L : List_Types) return List_Mal_Type;

   type List_Ptr is access all List_Mal_Type;

   function Deref_List (SP : Mal_Handle) return List_Ptr;


   type Lambda_Mal_Type is new Mal_Type with private;

   type Func_Type is (Prim_Binary, Mal_Func);
   function New_Lambda_Mal_Type
--     (Body : Mal_Handle := Smart_Pointers.Null_Smart_pointer)
       (Bin : Binary_Func_Access;
        Rep : Mal_String)
   return Mal_Handle;

   overriding function Sym_Type (T : Lambda_Mal_Type) return Sym_Types;

   -- primitive functions on Mal_Handle,
--   function "+" (A, B : Mal_Handle) return Mal_Handle;

   generic
      with function Int_Op (A, B : Mal_Integer) return Mal_Integer;
      with function Float_Op (A, B : Mal_Float) return Mal_Float;
   function Op (A, B : Mal_Handle) return Mal_Handle;

private

   type Mal_Type is abstract new Smart_Pointers.Base_Class with record
      Meta : Mal_Handle;
   end record;

   -- Not allowed to be abstract and private.  RM 3.9.3(10)
   -- So if you call this it'll just raise an exception.
   function To_Str (T : Mal_Type) return Mal_String; -- is abstract;

   type Int_Mal_Type is new Mal_Type with record
      Int_Val : Mal_Integer;
   end record;

   overriding function To_Str (T : Int_Mal_Type) return Mal_String;

   type Float_Mal_Type is new Mal_Type with record
      Float_Val : Mal_Float;
   end record;

   overriding function To_Str (T : Float_Mal_Type) return Mal_String;

   type Sym_Mal_Type is new Mal_Type with record
      Symbol : Character;
   end record;

   overriding function To_Str (T : Sym_Mal_Type) return Mal_String;

   type String_Mal_Type is new Mal_Type with record
      The_String : Ada.Strings.Unbounded.Unbounded_String;
   end record;

   overriding function To_Str (T : String_Mal_Type) return Mal_String;

   type Atom_Mal_Type is new Mal_Type with record
      The_Atom : Ada.Strings.Unbounded.Unbounded_String;
   end record;

   overriding function To_Str (T : Atom_Mal_Type) return Mal_String;

   type Error_Mal_Type is new Mal_Type with record
      Error_Msg : Ada.Strings.Unbounded.Unbounded_String;
   end record;

   overriding function To_Str (T : Error_Mal_Type) return Mal_String;

   type Unitary_Mal_Type is new Mal_Type with record
      The_Function : Unitary_Functions;
      The_Operand : Mal_Handle;
   end record;

   overriding function To_Str (T : Unitary_Mal_Type) return Mal_String;


   -- Nodes have to be a differnt type from a List;
   -- otherwise how do you represent a list within a list?
   type Node_Mal_Type is new Mal_Type with record
      Left, Right : Mal_Handle;
   end record;

   function New_Node_Mal_Type
     (Left, Right : Mal_Handle := Smart_Pointers.Null_Smart_pointer)
   return Mal_Handle;

   overriding function Sym_Type (T : Node_Mal_Type) return Sym_Types;

   procedure Append (To_List : in out Node_Mal_Type; Op : Mal_Handle);

   function Map_Nodes
    (Func_Ptr : Func_Access;
     L : Node_Mal_Type)
   return Mal_Handle;

   overriding function To_Str (T : Node_Mal_Type) return Mal_String;

   type Node_Ptr is access all Node_Mal_Type;

   function Deref_Node (SP : Mal_Handle) return Node_Ptr;


   type List_Mal_Type is new Mal_Type with record
      List_Type : List_Types;
      The_List : Mal_Handle;
   end record;

   overriding function To_Str (T : List_Mal_Type) return Mal_String;


   type Lambda_Mal_Type is new Mal_Type with record
       Bin : Binary_Func_Access;
       Rep : Ada.Strings.Unbounded.Unbounded_String;
   end record;

   overriding function To_Str (T : Lambda_Mal_Type) return Mal_String;


end Types;
