--------------------------------------------------------------------
--  This is a captive copy of a pure Ada package in the           --
--  Simple Components software collection, without the eponymous  --
--  crate's intricacies.                                          --
--  Source: http://www.dmitry-kazakov.de/ada/components.htm       --
--------------------------------------------------------------------

--                                                                    --
--  package IEEE_754                Copyright (c)  Dmitry A. Kazakov  --
--  Interface                                      Luebeck            --
--                                                 Summer, 2008       --
--                                                                    --
--                                Last revision :  11:26 27 Jul 2008  --
--                                                                    --
--  This  library  is  free software; you can redistribute it and/or  --
--  modify it under the terms of the GNU General Public  License  as  --
--  published by the Free Software Foundation; either version  2  of  --
--  the License, or (at your option) any later version. This library  --
--  is distributed in the hope that it will be useful,  but  WITHOUT  --
--  ANY   WARRANTY;   without   even   the   implied   warranty   of  --
--  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU  --
--  General  Public  License  for  more  details.  You  should  have  --
--  received  a  copy  of  the GNU General Public License along with  --
--  this library; if not, write to  the  Free  Software  Foundation,  --
--  Inc., 59 Temple Place - Suite 330, Boston, MA 02111-1307, USA.    --
--                                                                    --
--  As a special exception, if other files instantiate generics from  --
--  this unit, or you link this unit with other files to produce  an  --
--  executable, this unit does not by  itself  cause  the  resulting  --
--  executable to be covered by the GNU General Public License. This  --
--  exception  does not however invalidate any other reasons why the  --
--  executable file might be covered by the GNU Public License.       --
--____________________________________________________________________--

with Interfaces;

private package Excel_Out.IEEE_754 is

   subtype Byte is Interfaces.Unsigned_8;

   Not_A_Number_Error      : exception;
   Positive_Overflow_Error : exception;
   Negative_Overflow_Error : exception;

end Excel_Out.IEEE_754;
